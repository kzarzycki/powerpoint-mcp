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

      const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
        headless: false,
        ignoreHTTPSErrors: true,
        args: [
          '--disable-blink-features=AutomationControlled', // Avoid automation detection
        ],
      })

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

    // Handle one-time "Developer Mode" dialog if it appears
    try {
      const devModeDialog = page.getByText(/trust this add-in|developer mode|enable developer/i)
      await devModeDialog.waitFor({ state: 'visible', timeout: 10_000 })
      const confirmBtn = page.getByRole('button', { name: /ok|enable|trust|yes/i })
      await confirmBtn.click()
      console.log('[e2e] Developer mode dialog accepted')
    } catch {
      // No dialog = already accepted in this profile, continue
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
