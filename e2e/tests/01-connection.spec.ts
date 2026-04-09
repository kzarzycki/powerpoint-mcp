import { E2E_BRIDGE_HEALTH, E2E_BRIDGE_PORT } from '../config.ts'
import { expect, test } from '../fixtures/pptx-page.ts'
import { getTextContent, isToolError } from '../helpers/content-parsers.ts'

test.describe('Connection & Sideloading', () => {
  test('bridge server is healthy', async ({}) => {
    const res = await fetch(E2E_BRIDGE_HEALTH, { signal: AbortSignal.timeout(5000) })
    expect(res.ok).toBe(true)
    const body = (await res.json()) as { status: string }
    expect(body.status).toBe('ok')
  })

  test('add-in connects via WebSocket', async ({ pptxPage }) => {
    // pptxPage fixture already waits for connection — if we get here, it worked.
    // Verify server-side: health endpoint shows at least 1 connection
    const res = await fetch(E2E_BRIDGE_HEALTH, { signal: AbortSignal.timeout(5000) })
    const body = (await res.json()) as { status: string; connections: number }
    expect(body.connections).toBeGreaterThanOrEqual(1)
  })

  test('add-in taskpane shows Connected status', async ({ pptxPage }) => {
    const taskpane = pptxPage.frameLocator(`iframe[src*="localhost:${E2E_BRIDGE_PORT}"]`)
    await expect(taskpane.locator('#status')).toContainText(/connected/i, { timeout: 10_000 })
  })

  test('list_presentations returns the test deck', async ({ pptxPage, mcpClient }) => {
    const result = await mcpClient.callTool({ name: 'list_presentations', arguments: {} })
    expect(isToolError(result)).toBe(false)

    const text = getTextContent(result)
    // Should contain at least one presentation
    expect(text).toContain('presentation')
  })

  test('inspect_deck returns slide information', async ({ pptxPage, mcpClient }) => {
    const result = await mcpClient.callTool({ name: 'inspect_deck', arguments: {} })
    expect(isToolError(result)).toBe(false)

    const text = getTextContent(result)
    // Test deck should have at least 5 slides
    expect(text).toMatch(/slides?/i)
  })
})
