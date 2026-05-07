import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { type BrowserContext, test as base, chromium } from '@playwright/test'
import {
  ADDIN_CONNECT_TIMEOUT,
  BROWSER_PROFILE_DIR,
  buildSideloadUrl,
  E2E_BRIDGE_PORT,
  E2E_MCP_URL,
} from '../config.ts'
import { waitForAddinConnection } from '../helpers/wait-for-connection.ts'

/** Test-scoped fixtures (created per test) */
export interface PptxTestFixtures {
  pptxPage: Awaited<ReturnType<BrowserContext['newPage']>>
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
        headless: true,
        ignoreHTTPSErrors: true,
        userAgent:
          'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.7727.15 Safari/537.36',
        extraHTTPHeaders: {
          'sec-ch-ua': '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
          'sec-ch-ua-mobile': '?0',
          'sec-ch-ua-platform': '"macOS"',
        },
        args: ['--disable-blink-features=AutomationControlled'],
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

    // Handle "Enable Developer Mode" dialog if it appears.
    // The dialog has a checkbox that MUST be checked before clicking OK,
    // otherwise WAC ignores the confirmation and the add-in won't load.
    try {
      const devModeDialog = page.getByText(/trust this add-in|developer mode|enable developer/i)
      await devModeDialog.waitFor({ state: 'visible', timeout: 10_000 })
      console.log('[e2e] Developer mode dialog detected')
      // Check the "Enable Developer Mode now." checkbox.
      // Use locator().first() to avoid strict-mode errors when multiple checkboxes are present;
      // the unchecked dev-mode checkbox is the only one in this dialog.
      const enableCheckbox = page.locator('input[type="checkbox"]').first()
      await enableCheckbox.check()
      console.log('[e2e] Checkbox checked')
      const confirmBtn = page.getByRole('button', { name: /ok|enable|trust|yes/i })
      await confirmBtn.click()
      console.log('[e2e] Developer mode dialog accepted (checkbox checked)')
      // Wait for dialog to close and add-in to begin loading
      await page.waitForTimeout(2000)
    } catch (err) {
      console.log('[e2e] Dialog handling (skipped or errored):', err instanceof Error ? err.message : String(err))
    }

    // Wait for add-in taskpane iframe to appear
    const taskpaneIframe = page.frameLocator(`iframe[src*="localhost:${E2E_BRIDGE_PORT}"]`)
    try {
      await taskpaneIframe.locator('#status').waitFor({ state: 'visible', timeout: ADDIN_CONNECT_TIMEOUT })
    } catch {
      throw new Error(
        `Add-in taskpane iframe did not appear within ${ADDIN_CONNECT_TIMEOUT}ms. ` +
          'Office Web may not have processed the sideload URL parameters. ' +
          'Verify the document URL is correct and the bridge server is serving the manifest.',
      )
    }

    // Wait for WebSocket connection (server-side confirmation)
    await waitForAddinConnection()

    await use(page)
    await page.close()
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
