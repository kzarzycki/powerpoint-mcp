// ACP sidecar: bridges WebSocket frames to a stdio-spawned claude-agent-acp child.
// Intercepts session/new to inject the PowerPoint MCP server and a system-prompt append.
// Spike scope: 1 sidecar, per-connection child, plain ws (taskpane integration will use wss).

import { type ChildProcess, spawn } from 'node:child_process'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { WebSocketServer, type WebSocket } from 'ws'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const ACP_BIN = resolve(SCRIPT_DIR, 'node_modules/.bin/claude-agent-acp')

const PORT = Number(process.env.SIDECAR_PORT ?? 0) // 0 = OS-assigned
const HOST = process.env.SIDECAR_HOST ?? '127.0.0.1'
const MCP_URL = process.env.MCP_URL ?? 'http://127.0.0.1:3001/mcp'
const MCP_NAME = process.env.MCP_NAME ?? 'powerpoint'
const SYSTEM_PROMPT_APPEND =
  process.env.SYSTEM_PROMPT_APPEND ??
  'You can inspect and manipulate the user\'s open PowerPoint deck via the "powerpoint" MCP server. Prefer powerpoint MCP tools over guessing.'

type JsonRpcMessage = {
  jsonrpc?: '2.0'
  id?: string | number | null
  method?: string
  params?: Record<string, unknown>
  result?: unknown
  error?: unknown
}

function log(tag: string, line: string): void {
  const trimmed = line.length > 240 ? `${line.slice(0, 240)}…` : line
  process.stderr.write(`[sidecar:${tag}] ${trimmed}\n`)
}

function injectIntoNewSession(msg: JsonRpcMessage): JsonRpcMessage {
  if (msg.method !== 'session/new' || !msg.params) return msg
  const params = msg.params as {
    mcpServers?: unknown[]
    _meta?: Record<string, unknown>
  }

  const existing = Array.isArray(params.mcpServers) ? params.mcpServers : []
  const alreadyHasPowerPoint = existing.some(
    (s): s is { name: string } =>
      typeof s === 'object' && s !== null && (s as { name?: unknown }).name === MCP_NAME,
  )

  params.mcpServers = alreadyHasPowerPoint
    ? existing
    : [
        ...existing,
        {
          type: 'http',
          name: MCP_NAME,
          url: MCP_URL,
          headers: [],
        },
      ]

  params._meta = params._meta ?? {}
  if (!params._meta.systemPrompt) {
    params._meta.systemPrompt = {
      append: SYSTEM_PROMPT_APPEND,
      excludeDynamicSections: false,
    }
  }

  log('inject', `session/new augmented: +mcpServers[${MCP_NAME}] +_meta.systemPrompt`)
  return msg
}

function attachAgent(ws: WebSocket): void {
  const child: ChildProcess = spawn(ACP_BIN, [], {
    stdio: ['pipe', 'pipe', 'inherit'], // child stderr → our stderr (logs)
    env: { ...process.env },
  })
  log('spawn', `pid=${child.pid}`)

  if (!child.stdin || !child.stdout) {
    log('error', 'child stdio missing — killing')
    child.kill()
    ws.close(1011, 'agent stdio unavailable')
    return
  }

  // Agent → client (newline-delimited JSON)
  const rl = createInterface({ input: child.stdout, crlfDelay: Number.POSITIVE_INFINITY })
  rl.on('line', (line) => {
    const text = line.trim()
    if (!text) return
    log('a→c', text)
    if (ws.readyState === ws.OPEN) ws.send(text)
  })

  // Client → agent
  ws.on('message', (raw) => {
    const text = raw.toString()
    let outbound = text
    try {
      const parsed: JsonRpcMessage = JSON.parse(text)
      const augmented = injectIntoNewSession(parsed)
      if (augmented !== parsed || parsed.method === 'session/new') {
        outbound = JSON.stringify(augmented)
      }
    } catch (err) {
      log('warn', `non-JSON or parse error: ${(err as Error).message}`)
    }
    log('c→a', outbound)
    child.stdin?.write(`${outbound}\n`)
  })

  ws.on('close', (code, reason) => {
    log('ws-close', `code=${code} reason=${reason.toString()}`)
    child.kill('SIGTERM')
  })

  child.on('exit', (code, signal) => {
    log('agent-exit', `code=${code} signal=${signal}`)
    if (ws.readyState === ws.OPEN) ws.close(1011, 'agent exited')
  })
}

const wss = new WebSocketServer({ host: HOST, port: PORT })

wss.on('listening', () => {
  const addr = wss.address()
  if (typeof addr === 'object' && addr !== null) {
    const url = `ws://${HOST}:${addr.port}`
    log('listening', url)
    // Stdout: machine-readable handshake for the test client / launchers.
    process.stdout.write(`${JSON.stringify({ event: 'listening', url, port: addr.port })}\n`)
  }
})

wss.on('connection', (ws, req) => {
  log('connection', `from=${req.socket.remoteAddress}:${req.socket.remotePort}`)
  attachAgent(ws)
})

process.on('SIGINT', () => {
  log('signal', 'SIGINT — closing server')
  wss.close(() => process.exit(0))
})
