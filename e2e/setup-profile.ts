/**
 * One-time headed profile bootstrap for the e2e suite.
 *
 * Deletes any existing browser profile, starts the bridge, launches a headed
 * Chromium on a fresh persistent profile, shows a splash page asking the user to
 * sign in, then navigates to the test deck. From there everything is automated:
 * the sideload dialogs are accepted, the taskpane is opened from the ribbon if it
 * doesn't auto-show, and the script monitors the bridge until the add-in connects.
 * Once connected it confirms on the splash, closes, and the primed profile is
 * reused by `npm run test:e2e` in headless mode.
 *
 * Run whenever Entra cookies expire or you want a clean profile.
 */
import { type ChildProcess, spawn } from 'node:child_process'
import { existsSync, rmSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
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
import { connectAddin, type PptxPage } from './helpers/sideload.ts'

const PROJECT_ROOT = resolve(import.meta.dirname, '..')
const SIGN_IN_WINDOW_MS = 10 * 60_000

function splash(title: string, body: string, accent: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>
<style>
  html,body{height:100%;margin:0}
  body{display:flex;align-items:center;justify-content:center;
    font:16px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
    background:#faf9f8;color:#201f1e}
  .card{max-width:520px;padding:40px 48px;background:#fff;border-radius:12px;
    box-shadow:0 2px 16px rgba(0,0,0,.08);border-top:4px solid ${accent}}
  h1{margin:0 0 12px;font-size:22px}
  ol{margin:16px 0 0;padding-left:20px}
  li{margin:6px 0}
  .muted{color:#605e5c;font-size:14px;margin-top:20px}
</style></head><body><div class="card"><h1>${title}</h1>${body}</div></body></html>`
}

const LOGIN_SPLASH = splash(
  'PowerPoint MCP — e2e profile setup',
  `<p>This sets up a clean browser profile for the end-to-end tests.</p>
   <ol>
     <li>A Microsoft 365 sign-in page will open in a moment.</li>
     <li>Sign in with your account.</li>
     <li>Leave the rest to the script — it accepts the add-in dialogs and opens the taskpane automatically.</li>
   </ol>
   <p class="muted">This window monitors the bridge and closes itself once the add-in is connected. Don't close it manually.</p>`,
  '#d83b01',
)

const SUCCESS_SPLASH = splash(
  '✓ Add-in connected',
  `<p>The profile is primed and the add-in is talking to the bridge.</p>
   <p class="muted">Closing automatically… You can now run <code>npm run test:e2e</code>.</p>`,
  '#107c10',
)

const ERROR_SPLASH = splash(
  '✗ Setup did not complete',
  `<p>The add-in never connected. Most often this means sign-in wasn't finished in time.</p>
   <p class="muted">This window closes shortly. Re-run <code>npm run e2e:setup-profile</code> and sign in promptly.</p>`,
  '#a4262c',
)

async function pollHealth(url: string, label: string, timeoutMs: number): Promise<void> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(url, { signal: AbortSignal.timeout(2000) })
      if (res.ok && ((await res.json()) as { status: string }).status === 'ok') return
    } catch {}
    await new Promise((r) => setTimeout(r, HEALTH_POLL_INTERVAL))
  }
  throw new Error(`${label} did not become healthy within ${timeoutMs}ms at ${url}`)
}

/**
 * Wait for the PowerPoint Web deck to actually render — i.e. for the user to finish
 * the interactive M365 sign-in. The WAC host frame (euc-powerpoint.officeapps.live.com)
 * only appears once the document is loading, so poll for it. Logs progress so a slow
 * sign-in doesn't look like a hang.
 */
async function waitForDeck(page: PptxPage, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs
  let lastLog = 0
  while (Date.now() < deadline) {
    const loaded = page.frames().some((f) => {
      try {
        return new URL(f.url()).origin === 'https://euc-powerpoint.officeapps.live.com'
      } catch {
        return false
      }
    })
    if (loaded) return
    if (Date.now() - lastLog > 15_000) {
      console.log('[setup] Waiting for sign-in / deck to load…')
      lastLog = Date.now()
    }
    await page.waitForTimeout(2000)
  }
  throw new Error('Deck never loaded — was sign-in completed?')
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

  // Wipe the existing profile for a clean bootstrap.
  if (existsSync(BROWSER_PROFILE_DIR)) {
    console.log(`[setup] Deleting existing profile at ${BROWSER_PROFILE_DIR}`)
    rmSync(BROWSER_PROFILE_DIR, { recursive: true, force: true })
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

  const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
    headless: false,
    ignoreHTTPSErrors: true,
    userAgent:
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.7727.15 Safari/537.36',
    extraHTTPHeaders: {
      'sec-ch-ua': '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"macOS"',
    },
    args: [
      '--disable-blink-features=AutomationControlled',
      '--disable-session-crashed-bubble',
      '--hide-crash-restore-bubble',
    ],
  })

  try {
    await Promise.all([
      pollHealth(E2E_BRIDGE_HEALTH, 'Bridge', SERVER_START_TIMEOUT),
      pollHealth(E2E_MCP_HEALTH, 'MCP', SERVER_START_TIMEOUT),
    ])
    console.log(`[setup] Bridge ready on ${E2E_BRIDGE_URL}`)

    // Proxy the loopback bridge so the public WAC origin can reach it (Private
    // Network Access) — same shim the test fixture uses.
    for (const host of ['127.0.0.1', 'localhost']) {
      await context.route(`https://${host}:${E2E_BRIDGE_PORT}/**`, async (route) => {
        try {
          const response = await route.fetch()
          await route.fulfill({
            response,
            headers: {
              ...response.headers(),
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Private-Network': 'true',
            },
          })
        } catch {
          await route.continue()
        }
      })
    }

    const page = context.pages()[0] ?? (await context.newPage())

    // Show the splash, give the user a few seconds to read, then send them to the
    // deck — which redirects to M365 sign-in on a clean profile.
    await page.setContent(LOGIN_SPLASH)
    console.log('[setup] Splash shown. Sign in to Microsoft 365 in the browser window when it opens.')
    await page.waitForTimeout(4000)

    await page.goto(buildSideloadUrl(docUrl), { waitUntil: 'domcontentloaded', timeout: 120_000 })

    // Wait patiently for the user to finish signing in (the deck rendering is the
    // signal), then run the automated dialog + ribbon connect.
    console.log(`[setup] Sign in now. Waiting up to ${SIGN_IN_WINDOW_MS / 60_000} min for the deck...`)
    await waitForDeck(page, SIGN_IN_WINDOW_MS)
    console.log('[setup] Deck loaded. Accepting dialogs and connecting the add-in...')
    try {
      await connectAddin(page)
    } catch (err) {
      await page.setContent(ERROR_SPLASH).catch(() => {})
      await page.waitForTimeout(8000)
      throw err
    }

    console.log('[setup] Add-in connected. Profile primed.')
    await page.setContent(SUCCESS_SPLASH).catch(() => {})
    await page.waitForTimeout(3000)
  } finally {
    await context.close().catch(() => {})
    console.log('[setup] Stopping bridge server...')
    await stopBridge(serverProcess)
  }

  console.log(`[setup] Done. Profile primed at ${BROWSER_PROFILE_DIR}`)
  console.log('[setup] You can now run: npm run test:e2e')
}

main().catch((err) => {
  console.error('[setup] Error:', err instanceof Error ? err.message : err)
  process.exit(1)
})
