import { type ChildProcess, spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { resolve } from 'node:path'
import {
  E2E_BRIDGE_HEALTH,
  E2E_BRIDGE_PORT,
  E2E_BRIDGE_URL,
  E2E_MCP_HEALTH,
  E2E_MCP_PORT,
  HEALTH_POLL_INTERVAL,
  SERVER_START_TIMEOUT,
} from './config.ts'

const PROJECT_ROOT = resolve(import.meta.dirname, '..')

async function pollHealth(url: string, label: string, timeoutMs: number): Promise<void> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(url, { signal: AbortSignal.timeout(2000) })
      if (res.ok) {
        const body = (await res.json()) as { status: string }
        if (body.status === 'ok') return
      }
    } catch {
      // Server not ready yet
    }
    await new Promise((r) => setTimeout(r, HEALTH_POLL_INTERVAL))
  }
  throw new Error(`${label} did not become healthy within ${timeoutMs}ms at ${url}`)
}

export default async function globalSetup(): Promise<() => Promise<void>> {
  // Verify certs exist
  const certPath = resolve(PROJECT_ROOT, 'certs', 'localhost.pem')
  const keyPath = resolve(PROJECT_ROOT, 'certs', 'localhost-key.pem')
  if (!existsSync(certPath) || !existsSync(keyPath)) {
    throw new Error(`TLS certs not found at ${certPath}. Run: npm run setup-certs`)
  }

  // Verify .env has E2E_DOC_URL
  const docUrl = process.env.E2E_DOC_URL
  if (!docUrl) {
    throw new Error('E2E_DOC_URL not set. Copy e2e/.env.sample to e2e/.env and set the OneDrive document URL.')
  }

  // Check if ports are already in use (avoid conflict with dev instance)
  for (const [port, label] of [
    [E2E_BRIDGE_PORT, 'Bridge'],
    [E2E_MCP_PORT, 'MCP'],
  ] as const) {
    try {
      const res = await fetch(`https://localhost:${port}/health`, { signal: AbortSignal.timeout(1000) })
      if (res.ok) {
        throw new Error(`${label} port ${port} is already in use. Stop any running bridge/MCP server on that port.`)
      }
    } catch (err) {
      if (err instanceof Error && err.message.includes('already in use')) throw err
      // Connection refused = port is free, good
    }
  }

  // Start test bridge server
  const serverProcess: ChildProcess = spawn(
    'node',
    ['--experimental-strip-types', 'server/index.ts', '--http', '--bridge'],
    {
      cwd: PROJECT_ROOT,
      env: {
        ...process.env,
        BRIDGE_PORT: String(E2E_BRIDGE_PORT),
        MCP_PORT: String(E2E_MCP_PORT),
        BRIDGE_TLS: '1',
        NODE_TLS_REJECT_UNAUTHORIZED: '0', // Accept mkcert self-signed certs
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  )

  // Log server output for debugging
  serverProcess.stdout?.on('data', (data: Buffer) => {
    process.stderr.write(`[e2e-bridge] ${data.toString()}`)
  })
  serverProcess.stderr?.on('data', (data: Buffer) => {
    process.stderr.write(`[e2e-bridge] ${data.toString()}`)
  })

  // Wait for server process to not crash immediately
  const crashed = await Promise.race([
    new Promise<boolean>((resolve) => {
      serverProcess.on('exit', (code) => {
        process.stderr.write(`[e2e-bridge] Server exited with code ${code}\n`)
        resolve(true)
      })
    }),
    new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 2000)),
  ])
  if (crashed) {
    throw new Error('Bridge server crashed on startup. Check logs above.')
  }

  // Poll health endpoints
  await Promise.all([
    pollHealth(E2E_BRIDGE_HEALTH, 'Bridge server', SERVER_START_TIMEOUT),
    pollHealth(E2E_MCP_HEALTH, 'MCP server', SERVER_START_TIMEOUT),
  ])

  console.log(`[e2e] Bridge server ready on ${E2E_BRIDGE_URL} (MCP on port ${E2E_MCP_PORT})`)

  // Return teardown function
  return async () => {
    if (serverProcess.pid && !serverProcess.killed) {
      serverProcess.kill('SIGTERM')
      // Give it a moment to shut down gracefully
      await new Promise<void>((resolve) => {
        serverProcess.on('exit', () => resolve())
        setTimeout(() => {
          if (!serverProcess.killed) serverProcess.kill('SIGKILL')
          resolve()
        }, 3000)
      })
      console.log('[e2e] Bridge server stopped')
    }
  }
}
