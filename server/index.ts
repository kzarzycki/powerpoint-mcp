import { randomUUID } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { createServer as createHttpServer } from 'node:http'
import { createServer as createHttpsServer } from 'node:https'
import { homedir } from 'node:os'
import { dirname, extname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js'
import type { WebSocket } from 'ws'
import { WebSocketServer } from 'ws'

import { ConnectionPool } from './bridge.ts'
import { substituteManifestPort } from './manifest.ts'
import { clearSessionWarnings, registerTools } from './tools.ts'
import { runVersionCheck } from './version-check.ts'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BRIDGE_DEFAULT_HTTP_PORT = 8080
const BRIDGE_DEFAULT_HTTPS_PORT = 8443
const MCP_HTTP_PORT = Number(process.env.MCP_PORT) || 3001
const SCRIPT_DIR = typeof __dirname !== 'undefined' ? __dirname : dirname(fileURLToPath(import.meta.url))
const PROJECT_ROOT = resolve(SCRIPT_DIR, '..')
const BRIDGE_CERT_PATH = resolve(PROJECT_ROOT, 'certs', 'localhost.pem')
const BRIDGE_KEY_PATH = resolve(PROJECT_ROOT, 'certs', 'localhost-key.pem')
const ADDIN_STATIC_DIR = resolve(PROJECT_ROOT, 'addin')

// ---------------------------------------------------------------------------
// Flag parsing
// ---------------------------------------------------------------------------

const cliArgs = process.argv.slice(2)
const enableStdio = cliArgs.includes('--stdio')
const enableHttp = cliArgs.includes('--http')
const enableBridge = cliArgs.includes('--bridge')

// Default: --stdio if no flags given
const noFlagsGiven = !enableStdio && !enableHttp && !enableBridge
const stdioActive = enableStdio || noFlagsGiven
const httpActive = enableHttp
const bridgeActive = enableBridge

// ---------------------------------------------------------------------------
// Stderr redirect: when STDIO transport is active, all console output → stderr
// (stdout is reserved for MCP JSON-RPC messages)
// ---------------------------------------------------------------------------

if (stdioActive) {
  const stderrWrite = (...stderrArgs: Parameters<typeof console.log>) => {
    process.stderr.write(`${stderrArgs.map(String).join(' ')}\n`)
  }
  console.log = stderrWrite
  console.error = stderrWrite
}

// ---------------------------------------------------------------------------
// TTY guard: STDIO transport requires piped stdin
// ---------------------------------------------------------------------------

if (stdioActive && process.stdin.isTTY) {
  console.error(
    `Error: --stdio requires piped stdin (stdout is reserved for MCP JSON-RPC).

Usage:
  Pipe a client:    echo '...' | node server/index.ts
  Claude Desktop:   Configure in claude_desktop_config.json with --stdio --bridge
  HTTP mode:        node server/index.ts --http --bridge
  All transports:   node server/index.ts --stdio --http --bridge

Flags (composable):
  --stdio   STDIO MCP transport (stdin/stdout) — default if no flags given
  --http    HTTP MCP transport on port ${MCP_HTTP_PORT}
  --bridge  HTTP/WS bridge server for PowerPoint add-in connection`,
  )
  process.exit(1)
}

// ---------------------------------------------------------------------------
// TLS mode (config-driven via BRIDGE_TLS env var)
// ---------------------------------------------------------------------------

const bridgeTls = process.env.BRIDGE_TLS === '1'

if (bridgeActive && bridgeTls && (!existsSync(BRIDGE_CERT_PATH) || !existsSync(BRIDGE_KEY_PATH))) {
  console.error(
    'Error: BRIDGE_TLS=1 but TLS certificate files not found.\n' +
      `  Expected: ${BRIDGE_CERT_PATH} and ${BRIDGE_KEY_PATH}\n` +
      '  Run: npm run setup-certs',
  )
  process.exit(1)
}

const BRIDGE_PORT =
  Number(process.env.BRIDGE_PORT) || (bridgeTls ? BRIDGE_DEFAULT_HTTPS_PORT : BRIDGE_DEFAULT_HTTP_PORT)

// ---------------------------------------------------------------------------
// Auto-sideload add-in manifest
// ---------------------------------------------------------------------------

function autoSideloadManifest(tls: boolean, port: number): void {
  // Sideloading copies the add-in manifest into PowerPoint's sandboxed container,
  // which triggers a macOS TCC prompt ("node would like to access data from other
  // apps"). We use a versioned marker file (.sideloaded) to skip sideloading when
  // the version/port hasn't changed, so the prompt only appears on first install or
  // after an update. Use `npm run sideload` to force re-install.
  const markerFile = resolve(PROJECT_ROOT, '.sideloaded')
  const pkgPath = resolve(PROJECT_ROOT, 'package.json')

  let currentVersion = 'unknown'
  try {
    const pkg = JSON.parse(readFileSync(pkgPath, 'utf8'))
    currentVersion = pkg.version
  } catch {}

  const markerValue = `${currentVersion}:${port}`

  try {
    const existing = readFileSync(markerFile, 'utf8').trim()
    if (existing === markerValue) {
      console.error('[sideload] Add-in already installed (use `npm run sideload` to update)')
      return
    }
    console.error(`[sideload] Config changed (${existing} → ${markerValue}), re-sideloading`)
  } catch {
    // marker doesn't exist — first install
  }

  const defaultPort = tls ? BRIDGE_DEFAULT_HTTPS_PORT : BRIDGE_DEFAULT_HTTP_PORT
  const wefDir = join(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef')
  const manifestName = tls ? 'manifest-https.xml' : 'manifest.xml'
  const src = resolve(ADDIN_STATIC_DIR, manifestName)
  const dest = join(wefDir, 'manifest.xml')
  try {
    if (!existsSync(src)) return
    const template = readFileSync(src, 'utf8')
    const content = substituteManifestPort(template, defaultPort, port)
    mkdirSync(wefDir, { recursive: true })
    writeFileSync(dest, content)
    writeFileSync(markerFile, markerValue)
    console.error('[sideload] Add-in manifest installed for PowerPoint')
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err)
    console.error(`[sideload] Warning: could not sideload manifest: ${msg}`)
  }
}

// ---------------------------------------------------------------------------
// Shared state
// ---------------------------------------------------------------------------

const pool = new ConnectionPool()
const mcpHttpTransports = new Map<string, StreamableHTTPServerTransport>()

// ---------------------------------------------------------------------------
// Create an MCP server instance with tools registered
// ---------------------------------------------------------------------------

function createMcpServer(getSessionId: () => string | undefined, getActiveSessionCount: () => number): McpServer {
  const mcpServer = new McpServer({
    name: 'powerpoint-mcp',
    version: '0.1.0',
  })
  registerTools(mcpServer, pool, getSessionId, getActiveSessionCount)
  return mcpServer
}

// ---------------------------------------------------------------------------
// MIME type map (for bridge static file serving)
// ---------------------------------------------------------------------------

const MIME_TYPES = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.xml': 'application/xml; charset=UTF-8',
} as const

type KnownExt = keyof typeof MIME_TYPES

function getMimeType(filePath: string): string {
  const ext = extname(filePath) as KnownExt
  return MIME_TYPES[ext] ?? 'application/octet-stream'
}

// ---------------------------------------------------------------------------
// MCP HTTP transport helpers
// ---------------------------------------------------------------------------

function parseJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    req.on('data', (chunk: Buffer) => chunks.push(chunk))
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString()))
      } catch {
        reject(new Error('Invalid JSON body'))
      }
    })
    req.on('error', reject)
  })
}

function createMcpHttpSession(transport: StreamableHTTPServerTransport): McpServer {
  return createMcpServer(
    () => transport.sessionId ?? undefined,
    () => mcpHttpTransports.size,
  )
}

async function handleMcpPost(req: IncomingMessage, res: ServerResponse): Promise<void> {
  try {
    const body = await parseJsonBody(req)
    const sessionId = req.headers['mcp-session-id'] as string | undefined

    if (sessionId && mcpHttpTransports.has(sessionId)) {
      await mcpHttpTransports.get(sessionId)!.handleRequest(req, res, body)
    } else if (!sessionId && isInitializeRequest(body)) {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          mcpHttpTransports.set(sid, transport)
          console.error(`MCP HTTP session initialized: ${sid}`)
        },
      })
      transport.onclose = () => {
        const sid = transport.sessionId
        if (sid) {
          mcpHttpTransports.delete(sid)
          clearSessionWarnings(sid)
          console.error(`MCP HTTP session closed: ${sid}`)
        }
      }
      const mcpServer = createMcpHttpSession(transport)
      await mcpServer.connect(transport)
      await transport.handleRequest(req, res, body)
    } else {
      res.writeHead(400, { 'Content-Type': 'application/json' })
      res.end(
        JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Bad request: no valid session' }, id: null }),
      )
    }
  } catch (err) {
    console.error('MCP HTTP POST error:', err)
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32603, message: 'Internal error' }, id: null }))
    }
  }
}

async function handleMcpGet(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (!sessionId || !mcpHttpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' })
    res.end(
      JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }),
    )
    return
  }
  await mcpHttpTransports.get(sessionId)!.handleRequest(req, res)
}

async function handleMcpDelete(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined
  if (!sessionId || !mcpHttpTransports.has(sessionId)) {
    res.writeHead(400, { 'Content-Type': 'application/json' })
    res.end(
      JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Invalid or missing session ID' }, id: null }),
    )
    return
  }
  await mcpHttpTransports.get(sessionId)!.handleRequest(req, res)
}

// ---------------------------------------------------------------------------
// Static file handler (for bridge server)
// ---------------------------------------------------------------------------

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const rawUrl = (req.url ?? '/').split('?')[0]

  // Handle CORS preflight — including Chrome's Private Network Access (PNA) preflight.
  // WAC (officeapps.live.com) fetches the manifest from localhost, which is a public→loopback
  // request. Chrome sends an OPTIONS preflight with Access-Control-Request-Private-Network:true
  // and only proceeds if the response includes Access-Control-Allow-Private-Network:true.
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Allow-Private-Network': 'true',
    })
    res.end()
    return
  }

  if (rawUrl === '/health' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' })
    res.end(JSON.stringify({ status: 'ok', connections: pool.size }))
    return
  }

  if (rawUrl === '/api/test') {
    let target: ReturnType<ConnectionPool['resolveTarget']>
    try {
      target = pool.resolveTarget()
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err)
      res.writeHead(500, { 'Content-Type': 'application/json' })
      res.end(JSON.stringify({ error: message }))
      return
    }
    pool
      .sendCommand(
        'executeCode',
        {
          code: 'var c = context.presentation.slides.getCount(); await context.sync(); return c.value;',
        },
        target.ws,
      )
      .then((result) => {
        res.writeHead(200, { 'Content-Type': 'application/json' })
        res.end(JSON.stringify({ slideCount: result }))
      })
      .catch((err: Error) => {
        res.writeHead(500, { 'Content-Type': 'application/json' })
        res.end(JSON.stringify({ error: err.message }))
      })
    return
  }

  const urlPath = rawUrl === '/' ? '/index.html' : rawUrl
  console.error(`[bridge] GET ${rawUrl}`)
  const filePath = resolve(join(ADDIN_STATIC_DIR, urlPath))

  if (!filePath.startsWith(ADDIN_STATIC_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' })
    res.end('403 Forbidden')
    return
  }

  if (!existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('404 Not Found')
    return
  }

  const raw = readFileSync(filePath)
  const mimeType = getMimeType(filePath)
  // For XML manifests served over the bridge, substitute the default port with
  // the actual running port so the taskpane SourceLocation URLs are correct.
  const content =
    extname(filePath) === '.xml'
      ? substituteManifestPort(
          raw.toString(),
          bridgeTls ? BRIDGE_DEFAULT_HTTPS_PORT : BRIDGE_DEFAULT_HTTP_PORT,
          BRIDGE_PORT,
        )
      : raw
  res.writeHead(200, {
    'Content-Type': mimeType,
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Private-Network': 'true',
  })
  res.end(content)
}

// ---------------------------------------------------------------------------
// Start: Bridge server (HTTP or HTTPS based on BRIDGE_TLS)
// ---------------------------------------------------------------------------

if (bridgeActive) {
  autoSideloadManifest(bridgeTls, BRIDGE_PORT)

  const bridgeServer = bridgeTls
    ? createHttpsServer({ cert: readFileSync(BRIDGE_CERT_PATH), key: readFileSync(BRIDGE_KEY_PATH) }, serveStatic)
    : createHttpServer(serveStatic)

  const wss = new WebSocketServer({ server: bridgeServer })

  wss.on('connection', (ws: WebSocket) => {
    console.error(`[${new Date().toISOString()}] Add-in WebSocket connected`)

    ws.on('message', (data: Buffer) => {
      let msg: { type?: string; id?: string; data?: unknown; error?: { message?: string }; documentUrl?: string }
      try {
        msg = JSON.parse(data.toString())
      } catch {
        console.error('Invalid JSON from add-in:', data.toString())
        return
      }

      if ((msg.type === 'response' || msg.type === 'error') && msg.id) {
        pool.handleResponse(msg.id, msg.type, msg.data, msg.error?.message)
      }

      if (msg.type === 'ready') {
        const documentUrl = typeof msg.documentUrl === 'string' && msg.documentUrl.length > 0 ? msg.documentUrl : null
        const presentationId = pool.generateId(documentUrl)
        pool.add(presentationId, {
          ws,
          ready: true,
          presentationId,
          filePath: documentUrl,
        })
        console.error(`[${new Date().toISOString()}] Add-in ready: ${presentationId}`)
      }
    })

    ws.on('close', () => {
      const disconnectedId = pool.removeBySocket(ws)
      if (disconnectedId) {
        console.error(`[${new Date().toISOString()}] Add-in disconnected: ${disconnectedId}`)
      }
      pool.rejectPendingForSocket(ws)
    })

    ws.on('error', (err: Error) => {
      console.error('Add-in WebSocket error:', err.message)
    })
  })

  const bridgeScheme = bridgeTls ? 'https' : 'http'
  const bridgeWsScheme = bridgeTls ? 'wss' : 'ws'

  bridgeServer.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`Error: Port ${BRIDGE_PORT} is already in use. Another bridge instance may be running.`)
      process.exit(1)
    }
    throw err
  })

  bridgeServer.listen(BRIDGE_PORT, () => {
    console.error('Bridge server running')
    console.error(`  ${bridgeScheme.toUpperCase()}: ${bridgeScheme}://localhost:${BRIDGE_PORT}`)
    console.error(`  ${bridgeWsScheme.toUpperCase()}:  ${bridgeWsScheme}://localhost:${BRIDGE_PORT}`)
  })
}

// ---------------------------------------------------------------------------
// Start: HTTP MCP transport (on MCP_HTTP_PORT)
// ---------------------------------------------------------------------------

if (httpActive) {
  function handleMcpHttpRequest(req: IncomingMessage, res: ServerResponse): void {
    const url = (req.url ?? '/').split('?')[0]

    if (url === '/health' && req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'application/json' })
      res.end(
        JSON.stringify({
          status: 'ok',
          connections: pool.size,
          sessions: mcpHttpTransports.size,
          bridge: bridgeActive,
        }),
      )
      return
    }

    if (url === '/mcp') {
      if (req.method === 'POST') {
        handleMcpPost(req, res)
      } else if (req.method === 'GET') {
        handleMcpGet(req, res)
      } else if (req.method === 'DELETE') {
        handleMcpDelete(req, res)
      } else {
        res.writeHead(405)
        res.end()
      }
      return
    }

    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('MCP endpoint is at /mcp')
  }

  const mcpHttpServer = createHttpServer(handleMcpHttpRequest)
  mcpHttpServer.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`Error: Port ${MCP_HTTP_PORT} is already in use. Another instance may be running.`)
      process.exit(1)
    }
    throw err
  })
  mcpHttpServer.listen(MCP_HTTP_PORT, () => {
    console.error(`  MCP HTTP: http://localhost:${MCP_HTTP_PORT}/mcp`)
  })
}

// ---------------------------------------------------------------------------
// Start: STDIO MCP transport
// ---------------------------------------------------------------------------

if (stdioActive) {
  const stdioTransport = new StdioServerTransport()
  const stdioMcpServer = createMcpServer(
    () => 'stdio',
    () => 1,
  )
  stdioMcpServer.connect(stdioTransport).then(() => {
    console.error('MCP STDIO transport running')
  })
}

// ---------------------------------------------------------------------------
// Version check (non-blocking, fire-and-forget)
// ---------------------------------------------------------------------------

try {
  const pkg = JSON.parse(readFileSync(resolve(PROJECT_ROOT, 'package.json'), 'utf8'))
  runVersionCheck(pkg.version)
} catch {}

// ---------------------------------------------------------------------------
// Startup summary
// ---------------------------------------------------------------------------

const activeInterfaces = [
  stdioActive && 'STDIO',
  httpActive && `HTTP(:${MCP_HTTP_PORT})`,
  bridgeActive && `Bridge(:${BRIDGE_PORT})`,
]
  .filter(Boolean)
  .join(' + ')

console.error(`powerpoint-mcp started [${activeInterfaces}]`)
if (!bridgeActive) {
  console.error('  Note: Bridge not active. Tools requiring the add-in will return errors.')
}
