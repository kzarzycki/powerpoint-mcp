/**
 * One-time headed profile bootstrap for the e2e suite.
 *
 * Starts the bridge server, launches a headed Chromium via playwright-cli
 * pointing at the sideload URL, waits for the user to complete M365 SSO and
 * accept the "Enable developer mode" dialog, then closes the browser and
 * shuts the bridge down. The persistent profile is reused by `npm run test:e2e`
 * in headless mode.
 *
 * Run whenever Entra cookies expire.
 */
import { type ChildProcess, spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { resolve } from 'node:path'
import { stdin, stdout } from 'node:process'
import { createInterface } from 'node:readline/promises'
import {
  BROWSER_PROFILE_DIR,
  buildSideloadUrl,
  E2E_BRIDGE_HEALTH,
  E2E_BRIDGE_PORT,
  E2E_BRIDGE_URL,
  E2E_MCP_HEALTH,
  E2E_MCP_PORT,
  HEALTH_POLL_INTERVAL,
  SERVER_START_TIMEOUT,
} from './config.ts'
import { loadE2eEnv } from './helpers/load-env.ts'

const PROJECT_ROOT = resolve(import.meta.dirname, '..')
const PW_SESSION = 'e2e-pptx'

async function pollHealth(url: string, label: string, timeoutMs: number): Promise<void> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(url, { signal: AbortSignal.timeout(2000) })
      if (res.ok) {
        const body = (await res.json()) as { status: string }
        if (body.status === 'ok') return
      }
    } catch {}
    await new Promise((r) => setTimeout(r, HEALTH_POLL_INTERVAL))
  }
  throw new Error(`${label} did not become healthy within ${timeoutMs}ms at ${url}`)
}

async function stopBridge(proc: ChildProcess): Promise<void> {
  if (!proc.pid || proc.killed) return
  proc.kill('SIGTERM')
  await new Promise<void>((resolve) => {
    proc.on('exit', () => resolve())
    setTimeout(() => {
      if (!proc.killed) proc.kill('SIGKILL')
      resolve()
    }, 3000)
  })
}

function closeBrowserSession(): void {
  spawnSync('playwright-cli', ['-s', PW_SESSION, 'close'], { stdio: 'inherit' })
}

async function main(): Promise<void> {
  loadE2eEnv()

  const docUrl = process.env.E2E_DOC_URL
  if (!docUrl) {
    console.error('[setup] E2E_DOC_URL not set. Create e2e/local-config.json with {"E2E_DOC_URL": "<url>"}.')
    process.exit(1)
  }

  const certPath = resolve(PROJECT_ROOT, 'certs', 'localhost.pem')
  const keyPath = resolve(PROJECT_ROOT, 'certs', 'localhost-key.pem')
  if (!existsSync(certPath) || !existsSync(keyPath)) {
    console.error(`[setup] TLS certs not found at ${certPath}. Run: npm run setup-certs`)
    process.exit(1)
  }

  const pwCheck = spawnSync('playwright-cli', ['--version'], { stdio: 'pipe' })
  if (pwCheck.status !== 0) {
    console.error('[setup] playwright-cli not found. Install: npm install -g @playwright/cli@latest')
    process.exit(1)
  }

  console.log('[setup] Starting bridge server...')
  const serverProcess = spawn('node', ['--experimental-strip-types', 'server/index.ts', '--http', '--bridge'], {
    cwd: PROJECT_ROOT,
    env: {
      ...process.env,
      BRIDGE_PORT: String(E2E_BRIDGE_PORT),
      MCP_PORT: String(E2E_MCP_PORT),
      BRIDGE_TLS: '1',
      NODE_TLS_REJECT_UNAUTHORIZED: '0',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  serverProcess.stdout?.on('data', (d: Buffer) => process.stderr.write(`[bridge] ${d.toString()}`))
  serverProcess.stderr?.on('data', (d: Buffer) => process.stderr.write(`[bridge] ${d.toString()}`))

  try {
    await Promise.all([
      pollHealth(E2E_BRIDGE_HEALTH, 'Bridge', SERVER_START_TIMEOUT),
      pollHealth(E2E_MCP_HEALTH, 'MCP', SERVER_START_TIMEOUT),
    ])
    console.log(`[setup] Bridge ready on ${E2E_BRIDGE_URL}`)

    const sideloadUrl = buildSideloadUrl(docUrl)
    console.log('[setup] Launching headed Chromium via playwright-cli...')
    const openResult = spawnSync(
      'playwright-cli',
      ['-s', PW_SESSION, 'open', '--headed', '--persistent', `--profile=${BROWSER_PROFILE_DIR}`, sideloadUrl],
      { stdio: 'inherit' },
    )
    if (openResult.status !== 0) {
      throw new Error(`playwright-cli open failed with exit code ${openResult.status}`)
    }

    console.log('')
    console.log('[setup] A Chromium window is open at the test deck. Complete these steps:')
    console.log('         1. Sign in to Microsoft 365 if prompted')
    console.log('         2. Wait for the PowerPoint Web deck to load')
    console.log('         3. Accept the "Enable developer mode" dialog if it appears')
    console.log('         4. Confirm the add-in taskpane shows "Connected"')
    console.log('')
    const rl = createInterface({ input: stdin, output: stdout })
    await rl.question('[setup] Press Enter when done — the browser will close... ')
    rl.close()

    console.log('[setup] Closing browser session...')
    closeBrowserSession()
  } finally {
    console.log('[setup] Stopping bridge server...')
    await stopBridge(serverProcess)
  }

  console.log(`[setup] Done. Profile primed at ${BROWSER_PROFILE_DIR}`)
  console.log('[setup] You can now run: npm run test:e2e')
}

main().catch((err) => {
  console.error('[setup] Error:', err instanceof Error ? err.message : err)
  closeBrowserSession()
  process.exit(1)
})
