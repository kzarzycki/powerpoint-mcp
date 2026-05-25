import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { type BrowserContext, test as base, chromium } from '@playwright/test'
import { BROWSER_PROFILE_DIR, buildSideloadUrl, E2E_BRIDGE_PORT, E2E_MCP_URL } from '../config.ts'
import { type AddinFrame, connectAddin, findAddinFrame, type PptxPage } from '../helpers/sideload.ts'

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

    // Dismiss the sideload dialogs, then wait for the add-in to connect — opening
    // the pane from the ribbon if AutoShowTaskpaneWithDocument doesn't fire.
    await connectAddin(page)

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
