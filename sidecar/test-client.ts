// E2E spike: launches the sidecar, connects as ACP client, drives a session, asserts that
// claude-agent-acp invokes a tool from the powerpoint MCP server (proving ACP→MCP wiring).

import { type ChildProcess, spawn } from 'node:child_process'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import WebSocket from 'ws'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const SIDECAR_ENTRY = resolve(SCRIPT_DIR, 'index.ts')

const PROMPT_TEXT =
  process.env.SPIKE_PROMPT ??
  'List all currently open PowerPoint presentations using the powerpoint MCP server. Reply with the list (or say "none open" if there are none). Be brief.'

const TIMEOUT_MS = Number(process.env.SPIKE_TIMEOUT_MS ?? 120_000)

type JsonRpc = {
  jsonrpc: '2.0'
  id?: number
  method?: string
  params?: unknown
  result?: unknown
  error?: { code: number; message: string; data?: unknown }
}

function nowIso(): string {
  return new Date().toISOString().slice(11, 23)
}

function log(tag: string, msg: string): void {
  process.stderr.write(`[test:${nowIso()}:${tag}] ${msg}\n`)
}

function startSidecar(): Promise<{ child: ChildProcess; url: string }> {
  return new Promise((accept, reject) => {
    const child = spawn('node', ['--experimental-strip-types', SIDECAR_ENTRY], {
      stdio: ['ignore', 'pipe', 'inherit'],
      env: { ...process.env, SIDECAR_PORT: '0' },
    })

    const rl = createInterface({ input: child.stdout!, crlfDelay: Number.POSITIVE_INFINITY })
    const t = setTimeout(() => reject(new Error('sidecar did not announce listening')), 10_000)

    rl.on('line', (line) => {
      try {
        const evt = JSON.parse(line)
        if (evt.event === 'listening' && evt.url) {
          clearTimeout(t)
          accept({ child, url: evt.url })
        }
      } catch {
        // ignore
      }
    })
    child.on('exit', (code) => reject(new Error(`sidecar exited early code=${code}`)))
  })
}

class AcpClient {
  private nextId = 1
  private pending = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void }>()
  private updateLog: Array<{ method: string; params: unknown }> = []
  private ws: WebSocket

  constructor(ws: WebSocket) {
    this.ws = ws
    ws.on('message', (raw) => this.onFrame(raw.toString()))
  }

  private onFrame(text: string): void {
    let msg: JsonRpc
    try {
      msg = JSON.parse(text)
    } catch {
      log('frame-bad', text.slice(0, 200))
      return
    }

    // Notification or request from agent → us
    if (msg.method) {
      log('notify', `${msg.method}`)
      this.updateLog.push({ method: msg.method, params: msg.params })

      // Auto-allow permission requests (spike: no human in loop).
      if (msg.method === 'session/request_permission' && msg.id !== undefined) {
        const params = msg.params as { options?: Array<{ optionId: string; kind: string }> }
        const allow = params?.options?.find((o) => o.kind === 'allow_always' || o.kind === 'allow_once')
        const optionId = allow?.optionId ?? params?.options?.[0]?.optionId
        const reply: JsonRpc = {
          jsonrpc: '2.0',
          id: msg.id,
          result: { outcome: { outcome: 'selected', optionId } },
        }
        log('allow', `optionId=${optionId}`)
        this.ws.send(JSON.stringify(reply))
      }
      return
    }

    // Response to our request
    if (msg.id !== undefined && this.pending.has(msg.id as number)) {
      const slot = this.pending.get(msg.id as number)!
      this.pending.delete(msg.id as number)
      if (msg.error) slot.reject(new Error(`${msg.error.code} ${msg.error.message}`))
      else slot.resolve(msg.result)
    }
  }

  request<T>(method: string, params: unknown): Promise<T> {
    const id = this.nextId++
    const frame: JsonRpc = { jsonrpc: '2.0', id, method, params: params as Record<string, unknown> }
    log('send', method)
    this.ws.send(JSON.stringify(frame))
    return new Promise<T>((accept, reject) => {
      this.pending.set(id, { resolve: accept as (v: unknown) => void, reject })
    })
  }

  get updates(): Array<{ method: string; params: unknown }> {
    return this.updateLog
  }
}

async function main(): Promise<void> {
  log('boot', `starting sidecar from ${SIDECAR_ENTRY}`)
  const { child, url } = await startSidecar()
  log('boot', `sidecar up at ${url}, connecting WS`)

  const ws = new WebSocket(url)
  await new Promise<void>((accept, reject) => {
    ws.once('open', () => accept())
    ws.once('error', reject)
  })
  log('ws', 'connected')

  const client = new AcpClient(ws)

  const initRes = await client.request<unknown>('initialize', {
    protocolVersion: 1,
    clientInfo: { name: 'acp-spike-test-client', version: '0.0.0' },
  })
  log('init', JSON.stringify(initRes).slice(0, 300))

  const session = await client.request<{ sessionId: string }>('session/new', {
    cwd: SCRIPT_DIR,
    mcpServers: [], // sidecar will inject powerpoint MCP
  })
  log('session', JSON.stringify(session).slice(0, 300))

  // session/prompt is a long-running request that resolves when the agent stops.
  const promptDone = client.request<{ stopReason: string }>('session/prompt', {
    sessionId: session.sessionId,
    prompt: [{ type: 'text', text: PROMPT_TEXT }],
  })

  const overall = new Promise<{ stopReason: string }>((accept, reject) => {
    const t = setTimeout(() => reject(new Error('prompt timeout')), TIMEOUT_MS)
    promptDone.then(
      (v) => {
        clearTimeout(t)
        accept(v)
      },
      (e) => {
        clearTimeout(t)
        reject(e)
      },
    )
  })

  let result: { stopReason: string } | null = null
  let promptError: Error | null = null
  try {
    result = await overall
  } catch (e) {
    promptError = e as Error
  }

  // Assertions
  const updates = client.updates
  const sessionUpdates = updates.filter((u) => u.method === 'session/update')
  const toolCallUpdates = sessionUpdates.filter((u) => {
    const p = u.params as { update?: { sessionUpdate?: string; toolCall?: unknown } }
    return p?.update?.sessionUpdate === 'tool_call' || p?.update?.toolCall !== undefined
  })
  const powerpointToolCalls = toolCallUpdates.filter((u) => {
    const p = u.params as { update?: { title?: string; kind?: string; rawInput?: unknown } }
    const blob = JSON.stringify(p?.update ?? {})
    return /powerpoint|list_presentations|inspect_deck/i.test(blob)
  })

  log('summary', `total updates=${updates.length} session/update=${sessionUpdates.length} tool_calls=${toolCallUpdates.length} powerpoint_calls=${powerpointToolCalls.length}`)

  if (powerpointToolCalls.length > 0) {
    process.stderr.write('\n========================================\n')
    process.stderr.write('SPIKE PASS — agent invoked PowerPoint MCP\n')
    process.stderr.write('========================================\n')
    process.stderr.write(`Sample tool_call:\n${JSON.stringify(powerpointToolCalls[0], null, 2).slice(0, 800)}\n`)
    if (result) process.stderr.write(`stopReason: ${result.stopReason}\n`)
    ws.close()
    child.kill()
    process.exit(0)
  }

  process.stderr.write('\n========================================\n')
  process.stderr.write('SPIKE FAIL — no PowerPoint MCP tool call observed\n')
  process.stderr.write('========================================\n')
  if (promptError) process.stderr.write(`prompt error: ${promptError.message}\n`)
  if (result) process.stderr.write(`stopReason: ${result.stopReason}\n`)
  process.stderr.write(`Last 3 updates:\n${JSON.stringify(updates.slice(-3), null, 2).slice(0, 1500)}\n`)
  ws.close()
  child.kill()
  process.exit(1)
}

main().catch((err) => {
  process.stderr.write(`[test:fatal] ${(err as Error).stack ?? String(err)}\n`)
  process.exit(2)
})
