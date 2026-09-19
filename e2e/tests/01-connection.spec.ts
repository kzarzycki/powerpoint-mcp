import { E2E_BRIDGE_HEALTH } from '../config.ts'
import { expect, test } from '../fixtures/pptx-page.ts'
import { getJsonContent, getTextContent, isToolError } from '../helpers/content-parsers.ts'
import { fetchLoopbackJson } from '../helpers/loopback-fetch.ts'

interface DeckOverview {
  slideWidth: number
  slideHeight: number
  slides: Array<{ index: number; id: string; layout: string; shapeCount: number }>
}

test.describe('Connection & Sideloading', () => {
  test('bridge server is healthy', async ({}) => {
    const body = await fetchLoopbackJson<{ status: string }>(E2E_BRIDGE_HEALTH, 5000)
    expect(body.status).toBe('ok')
  })

  test('add-in connects via WebSocket', async ({ pptxPage }) => {
    // pptxPage fixture already waits for connection — if we get here, it worked.
    // Verify server-side: health endpoint shows at least 1 connection
    const body = await fetchLoopbackJson<{ status: string; connections: number }>(E2E_BRIDGE_HEALTH, 5000)
    expect(body.connections).toBeGreaterThanOrEqual(1)
  })

  test('add-in taskpane shows Connected status', async ({ addinFrame }) => {
    await expect(addinFrame.locator('#status')).toContainText(/connected/i, { timeout: 10_000 })
  })

  test('list_presentations returns the test deck', async ({ pptxPage, mcpClient }) => {
    const result = await mcpClient.callTool({ name: 'list_presentations', arguments: {} })
    expect(isToolError(result)).toBe(false)

    const text = getTextContent(result)
    // Should contain at least one presentation
    expect(text).toContain('presentation')
  })

  test('inspect_deck reports the 5-slide test deck', async ({ pptxPage, mcpClient }) => {
    const result = await mcpClient.callTool({ name: 'inspect_deck', arguments: {} })
    expect(isToolError(result)).toBe(false)

    const deck = getJsonContent<DeckOverview>(result)

    // The test deck has exactly 5 slides.
    expect(deck.slides).toHaveLength(5)

    // Slide dimensions are real, positive point values.
    expect(deck.slideWidth).toBeGreaterThan(0)
    expect(deck.slideHeight).toBeGreaterThan(0)

    // Each slide carries the expected structure: 0-based index, a stable id,
    // a layout name, and a shape count.
    deck.slides.forEach((slide, i) => {
      expect(slide.index).toBe(i)
      expect(slide.id).toBeTruthy()
      expect(slide.layout).toBeTruthy()
      expect(slide.shapeCount).toBeGreaterThanOrEqual(0)
    })
  })
})
