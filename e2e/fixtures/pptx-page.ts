import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { type BrowserContext, test as base, chromium } from '@playwright/test'
import { BROWSER_PROFILE_DIR, buildSideloadUrl, E2E_BRIDGE_PORT, E2E_MCP_URL } from '../config.ts'
import { waitForAddinConnection } from '../helpers/wait-for-connection.ts'

type PptxPage = Awaited<ReturnType<BrowserContext['newPage']>>
type AddinFrame = ReturnType<PptxPage['frames']>[number]

/**
 * Find the add-in taskpane frame. It loads from the bridge (localhost:PORT) but
 * is nested inside the WAC host frame, so a top-level locator/frameLocator can't
 * reach it — enumerate all frames instead.
 */
function findAddinFrame(page: PptxPage): AddinFrame | undefined {
  return page.frames().find((f) => f.url().includes(`localhost:${E2E_BRIDGE_PORT}/index.html`))
}

/**
 * Poll every frame of the page and dismiss the two sideload gating dialogs Office
 * Web raises before an add-in loads: "Enable Developer Mode" (check the box, then
 * OK) and "Registering Developer Add-in Manifest" (Yes). Returns once both are
 * handled, the taskpane frame shows up, or the window elapses.
 */
async function dismissSideloadDialogs(page: PptxPage, windowMs = 60_000): Promise<void> {
  const deadline = Date.now() + windowMs
  let devModeHandled = false
  let manifestHandled = false
  // Once the manifest dialog is accepted, keep polling only briefly — the pane
  // either auto-opens (addin frame appears) or it won't, and we fall back to the
  // ribbon. Avoids burning the whole window when auto-show doesn't fire.
  let settleDeadline = Number.POSITIVE_INFINITY
  while (Date.now() < deadline && Date.now() < settleDeadline) {
    // Taskpane already loaded — nothing left to dismiss.
    if (findAddinFrame(page)) return

    for (const frame of page.frames()) {
      try {
        if (
          !devModeHandled &&
          (await frame
            .getByText(/enable developer mode/i)
            .first()
            .isVisible())
        ) {
          await frame.locator('input[type="checkbox"]').first().check()
          await frame.getByRole('button', { name: /^ok$/i }).first().click()
          devModeHandled = true
          console.log('[e2e] Enabled developer mode')
          continue
        }
        if (
          !manifestHandled &&
          (await frame
            .getByText(/registering developer add-?in manifest/i)
            .first()
            .isVisible())
        ) {
          await frame.getByRole('button', { name: /^yes$/i }).first().click()
          manifestHandled = true
          settleDeadline = Date.now() + 10_000
          console.log('[e2e] Accepted add-in manifest registration')
        }
      } catch {
        // Frame detached or dialog vanished mid-check — keep polling.
      }
    }
    await page.waitForTimeout(1000)
  }
}

/**
 * Open the add-in taskpane from the ribbon. After a fresh manifest registration,
 * Office Web does NOT honor AutoShowTaskpaneWithDocument — the pane only auto-opens
 * once it has been opened manually at least once on the profile. The button lives
 * in the WAC host frame (euc-powerpoint.officeapps.live.com), labelled by the
 * manifest DisplayName. Returns true if a button was found and clicked.
 */
async function openTaskpaneFromRibbon(page: PptxPage, windowMs = 30_000): Promise<boolean> {
  const deadline = Date.now() + windowMs
  while (Date.now() < deadline) {
    for (const frame of page.frames()) {
      try {
        const button = frame.getByRole('button', { name: /^powerpoint mcp$/i }).first()
        if (await button.isVisible()) {
          await button.click()
          console.log('[e2e] Opened taskpane from ribbon')
          return true
        }
      } catch {
        // Frame detached or button not in this frame — keep scanning.
      }
    }
    await page.waitForTimeout(1000)
  }
  return false
}

/** Test-scoped fixtures (created per test) */
export interface PptxTestFixtures {
  pptxPage: PptxPage
  addinFrame: AddinFrame
  mcpClient: Client
}

/** Worker-scoped fixtures (created once per worker, shared across tests) */
export interface PptxWorkerFixtures {
  sharedContext: BrowserContext
}

/**
 * Playwright fixture that provides:
 * - pptxPage: a browser page with PowerPoint Web loaded and add-in connected
 * - mcpClient: a connected MCP client
 *
 * The browser uses a persistent profile (pre-logged-in to Microsoft 365).
 * The page navigates to the test presentation with sideload URL parameters.
 * The fixture waits for the add-in to connect before yielding.
 */
export const test = base.extend<PptxTestFixtures, PptxWorkerFixtures>({
  // Worker-scoped: persistent browser context (created once, reused across tests)
  sharedContext: [
    async ({}, use) => {
      const docUrl = process.env.E2E_DOC_URL
      if (!docUrl) throw new Error('E2E_DOC_URL not set')

      // WAC (Office Web) silently skips add-in sideloading when it detects "HeadlessChrome"
      // in the User-Agent or sec-ch-ua client hint headers. Spoof them to look like
      // regular Chrome so the sideload URL params are processed normally.
      const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
        headless: process.env.E2E_HEADED !== '1',
        ignoreHTTPSErrors: true,
        userAgent:
          'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.7727.15 Safari/537.36',
        extraHTTPHeaders: {
          'sec-ch-ua': '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
          'sec-ch-ua-mobile': '?0',
          'sec-ch-ua-platform': '"macOS"',
        },
        // Suppress the "Restore pages? Chromium didn't shut down correctly" bubble
        // that appears when a prior run was killed rather than closed cleanly.
        args: [
          '--disable-blink-features=AutomationControlled',
          '--disable-session-crashed-bubble',
          '--hide-crash-restore-bubble',
        ],
      })

      // Chrome's Private Network Access (PNA) policy blocks public HTTPS origins
      // (euc-powerpoint.officeapps.live.com) from accessing loopback addresses.
      // Playwright intercepts these requests at the CDP layer before PNA enforcement
      // and proxies them via its own Node.js fetch (which has no PNA restriction).
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

      await use(context)
      await context.close()
    },
    { scope: 'worker' },
  ],

  // Test-scoped: PowerPoint Web page with add-in connected
  pptxPage: async ({ sharedContext: context }, use) => {
    const docUrl = process.env.E2E_DOC_URL!
    const sideloadUrl = buildSideloadUrl(docUrl)

    const page = await context.newPage()
    await page.goto(sideloadUrl, { waitUntil: 'domcontentloaded', timeout: 60_000 })

    // Dismiss the sideload gating dialogs. Office Web shows up to two, in sequence:
    //   1. "Enable Developer Mode" — the checkbox MUST be checked before OK, or WAC
    //      ignores the confirmation. Persists per-profile once accepted.
    //   2. "Registering Developer Add-in Manifest" — "Yes" to load the local manifest.
    // They can take well over 10s to render and may live in a nested frame, so poll
    // every frame until both are handled or the taskpane iframe appears.
    await dismissSideloadDialogs(page)

    // On a primed profile the taskpane auto-opens (AutoShowTaskpaneWithDocument)
    // and connects back over WebSocket. On a freshly-registered profile auto-show
    // doesn't fire, so if the connection doesn't appear within a short grace window,
    // open the pane from the ribbon and wait for the full timeout. Gate readiness on
    // the server-side connection rather than a DOM probe — the taskpane frame is
    // nested in the WAC host frame and isn't reachable from a top-level frameLocator.
    try {
      await waitForAddinConnection(8_000)
    } catch {
      console.log('[e2e] Taskpane did not auto-open; opening from ribbon')
      if (!(await openTaskpaneFromRibbon(page))) {
        throw new Error('Add-in ribbon button never appeared; cannot open taskpane')
      }
      await waitForAddinConnection()
    }

    await use(page)
    await page.close()
  },

  // Test-scoped: the connected add-in taskpane frame (nested in the WAC host frame)
  addinFrame: async ({ pptxPage: page }, use) => {
    const frame = findAddinFrame(page)
    if (!frame) throw new Error('Add-in taskpane frame not found despite an active bridge connection')
    await use(frame)
  },

  // Test-scoped: MCP client
  mcpClient: async ({}, use) => {
    const transport = new StreamableHTTPClientTransport(new URL(`${E2E_MCP_URL}/mcp`))
    const client = new Client({ name: 'e2e-test', version: '1.0.0' })
    await client.connect(transport)

    await use(client)
    await client.close()
  },
})

export { expect } from '@playwright/test'
