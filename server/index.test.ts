import { type ChildProcessWithoutNullStreams, execFileSync, spawn } from 'node:child_process'
import { existsSync, mkdirSync } from 'node:fs'
import { request as httpRequest } from 'node:http'
import { request as httpsRequest } from 'node:https'
import { createServer, type Server } from 'node:net'
import { dirname, resolve } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import WebSocket from 'ws'

const MAX_HTTP_BODY_BYTES = 16 * 1024 * 1024
const STARTUP_TIMEOUT_MS = 10_000
const POLL_INTERVAL_MS = 50

type Health = {
  status: string
  connections: number
  sessions?: number
  bridge?: boolean
}

type HttpResult = {
  statusCode: number
  body: string
}

type WebSocketResult = {
  opened: boolean
  statusCode?: number
  socket?: WebSocket
}

let bridgePort: number
let mcpPort: number
let serverProcess: ChildProcessWithoutNullStreams

async function sleep(ms: number): Promise<void> {
  const { promise, resolve: resolveSleep } = Promise.withResolvers<void>()
  setTimeout(resolveSleep, ms)
  return promise
}

async function freePort(): Promise<number> {
  const { promise, resolve: resolvePort, reject } = Promise.withResolvers<number>()
  const server: Server = createServer()
  server.once('error', reject)
  server.listen(0, '127.0.0.1', () => {
    const address = server.address()
    if (!address || typeof address === 'string') {
      server.close()
      reject(new Error('Could not determine free port'))
      return
    }
    server.close((error) => {
      if (error) reject(error)
      else resolvePort(address.port)
    })
  })
  return promise
}

// Polls a real subprocess's real network readiness. Genuine integration-test
// exception: there is no in-process promise to await instead of the platform
// clock — the server under test is a separate OS process.
async function waitFor(
  check: () => Promise<boolean>,
  label: string,
  stderrGetter: () => string = () => '',
): Promise<void> {
  const deadline = Date.now() + STARTUP_TIMEOUT_MS
  let lastError: unknown
  while (Date.now() < deadline) {
    try {
      if (await check()) return
    } catch (error) {
      lastError = error
    }
    await sleep(POLL_INTERVAL_MS)
  }
  const detail = lastError instanceof Error ? lastError.message : String(lastError)
  throw new Error(`${label} did not become ready: ${detail}\nserver stderr:\n${stderrGetter()}`)
}

async function health(port: number, scheme: 'http' | 'https' = 'http'): Promise<Health> {
  const requestFn = scheme === 'https' ? httpsRequest : httpRequest
  const { promise, resolve: resolveHealth, reject: rejectHealth } = Promise.withResolvers<Health>()
  const req = requestFn(
    { hostname: '127.0.0.1', port, path: '/health', method: 'GET', rejectUnauthorized: false },
    (res) => {
      const chunks: Buffer[] = []
      res.on('data', (chunk: Buffer) => chunks.push(chunk))
      res.on('end', () => {
        try {
          resolveHealth(JSON.parse(Buffer.concat(chunks).toString()) as Health)
        } catch (error) {
          rejectHealth(error instanceof Error ? error : new Error(String(error)))
        }
      })
    },
  )
  req.on('error', rejectHealth)
  req.end()
  return promise
}

const PROJECT_ROOT = resolve(import.meta.dirname, '..')
const TLS_CERT_PATH = resolve(PROJECT_ROOT, 'certs', 'localhost.pem')
const TLS_KEY_PATH = resolve(PROJECT_ROOT, 'certs', 'localhost-key.pem')

// The bridge's HTTPS mode reads a fixed certs/ path (see server/index.ts).
// CI has no mkcert-generated dev certs, so generate a throwaway self-signed
// pair via the platform's own openssl (present on macOS and GitHub's Ubuntu
// runners) the first time this suite needs one. Never touches a real cert
// `npm run setup-certs` already created.
function ensureTlsCertsExist(): boolean {
  if (existsSync(TLS_CERT_PATH) && existsSync(TLS_KEY_PATH)) return true
  try {
    execFileSync('openssl', ['version'], { stdio: 'ignore' })
  } catch {
    return false
  }
  mkdirSync(dirname(TLS_CERT_PATH), { recursive: true })
  execFileSync('openssl', [
    'req',
    '-x509',
    '-newkey',
    'rsa:2048',
    '-nodes',
    '-keyout',
    TLS_KEY_PATH,
    '-out',
    TLS_CERT_PATH,
    '-days',
    '1',
    '-subj',
    '/CN=localhost',
    '-addext',
    'subjectAltName=DNS:localhost,IP:127.0.0.1',
  ])
  return true
}

async function postBody(body: string, chunkSize?: number): Promise<HttpResult> {
  const { promise, resolve: resolveResult, reject: rejectResult } = Promise.withResolvers<HttpResult>()
  const req = httpRequest(
    {
      hostname: '127.0.0.1',
      port: mcpPort,
      path: '/mcp',
      method: 'POST',
      headers: {
        Accept: 'application/json, text/event-stream',
        'Content-Type': 'application/json',
        ...(chunkSize ? {} : { 'Content-Length': Buffer.byteLength(body) }),
      },
    },
    (res) => {
      const chunks: Buffer[] = []
      res.on('data', (chunk: Buffer) => chunks.push(chunk))
      res.on('end', () => {
        resolveResult({ statusCode: res.statusCode ?? 0, body: Buffer.concat(chunks).toString('utf8') })
      })
    },
  )
  req.on('error', rejectResult)
  if (chunkSize) {
    for (let offset = 0; offset < body.length; offset += chunkSize) {
      req.write(body.slice(offset, offset + chunkSize))
    }
    req.end()
  } else {
    req.end(body)
  }
  return promise
}

// `scheme`/`origin`/`host` are independent of `port`'s own scheme so a caller
// can point a plain `ws://` socket at a TLS bridge's port to prove rejection,
// not just point it at its own matching scheme.
async function openWebSocket(
  port: number,
  options: { scheme?: 'ws' | 'wss'; origin?: string; host?: string } = {},
): Promise<WebSocketResult> {
  const scheme = options.scheme ?? 'ws'
  const host = options.host ?? `127.0.0.1:${port}`
  const { promise, resolve: resolveSocket } = Promise.withResolvers<WebSocketResult>()
  const socket = new WebSocket(`${scheme}://127.0.0.1:${port}`, {
    rejectUnauthorized: false,
    headers: {
      Host: host,
      ...(options.origin ? { Origin: options.origin } : {}),
    },
  })
  let settled = false
  const finish = (result: WebSocketResult) => {
    if (settled) return
    settled = true
    if (!result.opened) socket.terminate()
    resolveSocket(result)
  }
  socket.once('open', () => finish({ opened: true, socket }))
  socket.once('unexpected-response', (_request, response) => finish({ opened: false, statusCode: response.statusCode }))
  socket.once('error', () => finish({ opened: false }))
  socket.once('close', () => finish({ opened: false }))
  return promise
}

function initializeBody(targetBytes: number): string {
  // The SDK validates the full envelope against a strict JSON-RPC request
  // schema, so padding must live inside `params` (which allows extra keys),
  // not as a sibling of `jsonrpc`/`method` (which would fail that check).
  const params = {
    protocolVersion: '2025-06-18',
    capabilities: {},
    clientInfo: { name: 'index-test', version: '1.0.0' },
    padding: '',
  }
  const emptyBody = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params })
  const paddingLength = Math.max(0, targetBytes - Buffer.byteLength(emptyBody))
  params.padding = 'x'.repeat(paddingLength)
  return JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params })
}

async function disconnectMidUpload(): Promise<void> {
  const { promise, resolve: resolveDisconnect } = Promise.withResolvers<void>()
  const req = httpRequest({
    hostname: '127.0.0.1',
    port: mcpPort,
    path: '/mcp',
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Transfer-Encoding': 'chunked' },
  })
  req.once('error', () => resolveDisconnect())
  req.write('{"jsonrpc":"2.0","method":"initialize","params":{"padding":"')
  req.destroy()
  return promise
}

async function spawnBridge(
  env: Record<string, string>,
  extraArgs: string[] = [],
): Promise<{ proc: ChildProcessWithoutNullStreams; stderr: () => string }> {
  let stderr = ''
  const proc = spawn(
    process.execPath,
    ['--experimental-strip-types', resolve(process.cwd(), 'server/index.ts'), '--bridge', ...extraArgs],
    {
      cwd: process.cwd(),
      env: { ...process.env, BRIDGE_NO_UPDATE_CHECK: '1', ...env },
      stdio: 'pipe',
    },
  )
  proc.stderr.on('data', (chunk: Buffer) => {
    stderr += chunk.toString()
  })
  return { proc, stderr: () => stderr }
}

async function killProcess(proc: ChildProcessWithoutNullStreams): Promise<void> {
  if (proc.killed) return
  const { promise, resolve: resolveExit } = Promise.withResolvers<void>()
  proc.once('exit', () => resolveExit())
  proc.kill('SIGTERM')
  // Real process teardown, not asserted test behavior — bound the wait so a
  // stuck subprocess can't hang the suite.
  await Promise.race([promise, sleep(2_000)])
}

let serverStderrGetter: () => string = () => ''

beforeAll(async () => {
  bridgePort = await freePort()
  mcpPort = await freePort()
  const spawned = await spawnBridge({ BRIDGE_PORT: String(bridgePort), MCP_PORT: String(mcpPort) }, ['--http'])
  serverProcess = spawned.proc
  serverStderrGetter = spawned.stderr
  await waitFor(
    async () => {
      const response = await fetch(`http://127.0.0.1:${mcpPort}/health`)
      return response.ok
    },
    'MCP health endpoint',
    serverStderrGetter,
  )
})

afterAll(async () => {
  await killProcess(serverProcess)
})

describe('bridge WebSocket origin policy', () => {
  it('registers the manifest localhost origin and rejects a foreign origin before ready', async () => {
    const allowed = await openWebSocket(bridgePort, {
      origin: `http://localhost:${bridgePort}`,
      host: `localhost:${bridgePort}`,
    })
    expect(allowed.opened).toBe(true)
    allowed.socket!.send(JSON.stringify({ type: 'ready', documentUrl: 'origin-test.pptx' }))
    await waitFor(
      async () => (await health(bridgePort)).connections === 1,
      'allowed bridge connection',
      serverStderrGetter,
    )

    const foreign = await openWebSocket(bridgePort, {
      origin: 'https://foreign.example.test',
      host: `localhost:${bridgePort}`,
    })
    expect(foreign.opened).toBe(false)
    expect((await health(bridgePort)).connections).toBe(1)

    allowed.socket!.close()
    await waitFor(
      async () => (await health(bridgePort)).connections === 0,
      'allowed bridge disconnection',
      serverStderrGetter,
    )
  })

  it('rejects a missing Origin header', async () => {
    const result = await openWebSocket(bridgePort, { host: `localhost:${bridgePort}` })
    expect(result.opened).toBe(false)
    expect((await health(bridgePort)).connections).toBe(0)
  })
})

const tlsCertsAvailable = ensureTlsCertsExist()

// CI runners without openssl (or without permission to write certs/) skip
// this suite rather than hang; local dev machines always have openssl.
describe.skipIf(!tlsCertsAvailable)('bridge WebSocket origin policy — TLS mode (PowerPoint Web path)', () => {
  let tlsBridgePort: number
  let tlsProcess: ChildProcessWithoutNullStreams
  let tlsStderrGetter: () => string = () => ''

  beforeAll(async () => {
    tlsBridgePort = await freePort()
    const tlsMcpPort = await freePort()
    const spawned = await spawnBridge({
      BRIDGE_TLS: '1',
      BRIDGE_PORT: String(tlsBridgePort),
      MCP_PORT: String(tlsMcpPort),
    })
    tlsProcess = spawned.proc
    tlsStderrGetter = spawned.stderr
    await waitFor(
      async () => (await health(tlsBridgePort, 'https')).status === 'ok',
      'TLS bridge health endpoint',
      tlsStderrGetter,
    )
  })

  afterAll(async () => {
    await killProcess(tlsProcess)
  })

  it('accepts the bridge\u2019s own https origin (what the PowerPoint Web taskpane iframe navigates to) and rejects a foreign origin', async () => {
    const allowed = await openWebSocket(tlsBridgePort, {
      scheme: 'wss',
      origin: `https://localhost:${tlsBridgePort}`,
      host: `localhost:${tlsBridgePort}`,
    })
    expect(allowed.opened).toBe(true)
    allowed.socket!.send(JSON.stringify({ type: 'ready', documentUrl: 'tls-origin-test.pptx' }))
    await waitFor(
      async () => (await health(tlsBridgePort, 'https')).connections === 1,
      'TLS bridge connection',
      tlsStderrGetter,
    )

    const foreign = await openWebSocket(tlsBridgePort, {
      scheme: 'wss',
      origin: 'https://foreign.example.test',
      host: `localhost:${tlsBridgePort}`,
    })
    expect(foreign.opened).toBe(false)
    expect((await health(tlsBridgePort, 'https')).connections).toBe(1)

    allowed.socket!.close()
    await waitFor(
      async () => (await health(tlsBridgePort, 'https')).connections === 0,
      'TLS bridge disconnection',
      tlsStderrGetter,
    )
  })
})

describe('bounded MCP HTTP request bodies', () => {
  it('accepts a legitimate JSON request near the byte limit', async () => {
    const result = await postBody(initializeBody(MAX_HTTP_BODY_BYTES - 1024))
    expect(result.statusCode).toBe(200)
  })

  it('applies the byte limit to chunked bodies', async () => {
    const body = initializeBody(MAX_HTTP_BODY_BYTES - 1024)
    const result = await postBody(body, 1024)
    expect(result.statusCode).toBe(200)
  })

  it('rejects oversized bodies without affecting a later request', async () => {
    const result = await postBody(initializeBody(MAX_HTTP_BODY_BYTES + 1))
    expect(result.statusCode).toBe(413)
    expect(result.body).toContain('Request body too large')

    const check = await fetch(`http://127.0.0.1:${mcpPort}/health`)
    expect(check.status).toBe(200)
  })

  it('rejects an oversized chunked body and remains responsive', async () => {
    const result = await postBody(initializeBody(MAX_HTTP_BODY_BYTES + 1), 4096)
    expect(result.statusCode).toBe(413)
    const check = await fetch(`http://127.0.0.1:${mcpPort}/health`)
    expect(check.status).toBe(200)
  })

  it('settles a mid-upload disconnect and accepts a subsequent request', async () => {
    await disconnectMidUpload()
    const result = await postBody(JSON.stringify({ jsonrpc: '2.0', method: 'not-initialize', id: 1 }))
    expect(result.statusCode).toBe(400)
  })
})
