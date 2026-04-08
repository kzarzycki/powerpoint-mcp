import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import JSZip from 'jszip'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import { ConnectionPool } from './bridge.ts'
import { localCopyCache, parseSlideRange, registerTools } from './tools.ts'

vi.mock('node:fs', () => ({ existsSync: vi.fn(() => true), readFileSync: vi.fn(), writeFileSync: vi.fn() }))

const SAMPLE_SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r></a:p>
          <a:p><a:r><a:rPr lang="en-US"/><a:t>World</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Content 2"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Body text</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`

async function makeSlideZipBase64(xmlContent?: string): Promise<string> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', xmlContent ?? SAMPLE_SLIDE_XML)
  return await zip.generateAsync({ type: 'base64' })
}

function mockWs(): WebSocket {
  return { send: vi.fn(), readyState: 1 } as unknown as WebSocket
}

async function setupMcpClient(pool: ConnectionPool) {
  const server = new McpServer({ name: 'test', version: '0.0.1' })
  registerTools(
    server,
    pool,
    () => 'test-session',
    () => 1,
  )

  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  await server.connect(serverTransport)

  const client = new Client({ name: 'test-client', version: '0.0.1' })
  await client.connect(clientTransport)

  return { client, server }
}

describe('MCP Tools', () => {
  let pool: ConnectionPool

  beforeEach(() => {
    pool = new ConnectionPool(100)
  })

  it('lists all 23 tools', async () => {
    const { client } = await setupMcpClient(pool)
    const result = await client.listTools()
    const names = result.tools.map((t) => t.name).sort()
    expect(names).toEqual([
      'copy_slides',
      'duplicate_slide',
      'edit_shape_paragraphs',
      'edit_slide_chart',
      'edit_slide_xml',
      'edit_slide_zip',
      'execute_officejs',
      'format_shapes',
      'get_local_copy',
      'insert_image',
      'inspect_deck',
      'inspect_layouts',
      'inspect_slide',
      'list_presentations',
      'preview_deck',
      'read_deck_text',
      'read_shape_paragraphs',
      'read_slide_xml',
      'read_slide_zip',
      'scan_slide',
      'screenshot_slide',
      'search_fluent_icons',
      'search_text',
      'verify_slides',
    ])
  })

  describe('list_presentations', () => {
    it('returns empty message with no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'list_presentations', arguments: {} })
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })

    it('returns connection info when presentations are connected', async () => {
      const ws = mockWs()
      pool.add('demo.pptx', {
        ws,
        ready: true,
        presentationId: 'demo.pptx',
        filePath: '/path/demo.pptx',
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'list_presentations', arguments: {} })
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed).toHaveLength(1)
      expect(parsed[0].presentationId).toBe('demo.pptx')
      expect(parsed[0].ready).toBe(true)
    })
  })

  describe('inspect_deck', () => {
    it('returns error with no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'inspect_deck', arguments: {} })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('screenshot_slide', () => {
    it('returns image content block on success', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'screenshot_slide',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('getImageAsBase64')
      expect(sentJson.params.code).toContain('width: 720')

      pool.handleResponse(sentJson.id, 'response', {
        base64: 'iVBORw0KGgo=',
        slideIndex: 0,
        slideId: 'slide-abc',
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; data?: string; mimeType?: string; text?: string }>

      expect(content[0].type).toBe('image')
      expect(content[0].data).toBe('iVBORw0KGgo=')
      expect(content[0].mimeType).toBe('image/png')

      expect(content[1].type).toBe('text')
      expect(content[1].text).toContain('Slide 0')
      expect(content[1].text).toContain('slide-abc')
    })

    it('passes custom width and height to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'screenshot_slide',
        arguments: { slideIndex: 0, width: 1280, height: 720 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('width: 1280')
      expect(sentJson.params.code).toContain('height: 720')

      pool.handleResponse(sentJson.id, 'response', {
        base64: 'abc123',
        slideIndex: 0,
        slideId: 'slide-xyz',
      })

      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'screenshot_slide',
        arguments: { slideIndex: 0 },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('copy_slides', () => {
    it('exports from source and inserts into destination', async () => {
      const sourceWs = mockWs()
      const destWs = mockWs()
      pool.add('source.pptx', {
        ws: sourceWs,
        ready: true,
        presentationId: 'source.pptx',
        filePath: '/path/source.pptx',
      })
      pool.add('dest.pptx', {
        ws: destWs,
        ready: true,
        presentationId: 'dest.pptx',
        filePath: '/path/dest.pptx',
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 2,
          sourcePresentationId: 'source.pptx',
          destinationPresentationId: 'dest.pptx',
        },
      })

      // Wait for export command to be sent to source
      await new Promise((r) => setTimeout(r, 10))

      const exportJson = JSON.parse((sourceWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(exportJson.action).toBe('executeCode')
      expect(exportJson.params.code).toContain('exportAsBase64')

      // Respond with exported Base64
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'UEsDBBQ=',
        slideIndex: 2,
        slideId: 'slide-src',
      })

      // Wait for insert command to be sent to destination
      await new Promise((r) => setTimeout(r, 10))

      const insertJson = JSON.parse((destWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(insertJson.action).toBe('executeCode')
      expect(insertJson.params.code).toContain('insertSlidesFromBase64')
      expect(insertJson.params.code).toContain('UEsDBBQ=')

      // Respond with insert result
      pool.handleResponse(insertJson.id, 'response', { slideCount: 8 })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.copied.slideIndex).toBe(2)
      expect(parsed.copied.slideId).toBe('slide-src')
      expect(parsed.destination.slideCount).toBe(8)
    })

    it('passes formatting and targetSlideId options', async () => {
      const sourceWs = mockWs()
      const destWs = mockWs()
      pool.add('a.pptx', {
        ws: sourceWs,
        ready: true,
        presentationId: 'a.pptx',
        filePath: null,
      })
      pool.add('b.pptx', {
        ws: destWs,
        ready: true,
        presentationId: 'b.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'a.pptx',
          destinationPresentationId: 'b.pptx',
          formatting: 'UseDestinationTheme',
          targetSlideId: '267#',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const exportJson = JSON.parse((sourceWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'DATA',
        slideIndex: 0,
        slideId: 'slide-0',
      })

      await new Promise((r) => setTimeout(r, 10))

      const insertJson = JSON.parse((destWs.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(insertJson.params.code).toContain('UseDestinationTheme')
      expect(insertJson.params.code).toContain('267#')

      pool.handleResponse(insertJson.id, 'response', { slideCount: 4 })
      await toolPromise
    })

    it('returns error when source presentation not found', async () => {
      const ws = mockWs()
      pool.add('dest.pptx', {
        ws,
        ready: true,
        presentationId: 'dest.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'missing.pptx',
          destinationPresentationId: 'dest.pptx',
        },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('missing.pptx')
    })

    it('returns error when destination presentation not found', async () => {
      const ws = mockWs()
      pool.add('source.pptx', {
        ws,
        ready: true,
        presentationId: 'source.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'copy_slides',
        arguments: {
          sourceSlideIndex: 0,
          sourcePresentationId: 'source.pptx',
          destinationPresentationId: 'missing.pptx',
        },
      })

      // Export succeeds
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', {
        base64: 'DATA',
        slideIndex: 0,
        slideId: 'slide-0',
      })

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('missing.pptx')
    })
  })

  describe('insert_image', () => {
    it('passes base64 data directly into the code string', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'iVBORw0KGgoAAAANSUhEUg==',
          sourceType: 'base64',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      expect(sentJson.params.code).toContain('iVBORw0KGgoAAAANSUhEUg==')
      expect(sentJson.params.code).toContain('CoercionType.Image')

      pool.handleResponse(sentJson.id, 'response', { success: true })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.success).toBe(true)
    })

    it('reads file and base64 encodes it', async () => {
      const { readFileSync } = await import('node:fs')
      vi.mocked(readFileSync).mockReturnValue(Buffer.from([0x89, 0x50, 0x4e, 0x47]))

      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: '/path/to/image.png',
          sourceType: 'file',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      expect(readFileSync).toHaveBeenCalledWith('/path/to/image.png')

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      // The base64 of [0x89, 0x50, 0x4e, 0x47] is "iVBORw=="
      expect(sentJson.params.code).toContain(Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64'))

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('fetches URL and base64 encodes it', async () => {
      const mockArrayBuffer = new Uint8Array([1, 2, 3]).buffer
      globalThis.fetch = vi.fn().mockResolvedValue({
        ok: true,
        arrayBuffer: () => Promise.resolve(mockArrayBuffer),
      })

      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'https://example.com/image.png',
          sourceType: 'url',
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      expect(globalThis.fetch).toHaveBeenCalledWith('https://example.com/image.png')

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('setSelectedDataAsync')
      expect(sentJson.params.code).toContain(Buffer.from(new Uint8Array([1, 2, 3])).toString('base64'))

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('wraps with goToByIdAsync when slideIndex is provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'AAAA',
          sourceType: 'base64',
          slideIndex: 2,
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      // slideIndex 2 (0-based) → goToByIdAsync(3, ...) (1-based)
      expect(sentJson.params.code).toContain('goToByIdAsync(3,')
      expect(sentJson.params.code).toContain('GoToType.Index')

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('includes positioning options when provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'BBBB',
          sourceType: 'base64',
          left: 100,
          top: 50,
          width: 400,
          height: 300,
        },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('imageLeft: 100')
      expect(sentJson.params.code).toContain('imageTop: 50')
      expect(sentJson.params.code).toContain('imageWidth: 400')
      expect(sentJson.params.code).toContain('imageHeight: 300')

      pool.handleResponse(sentJson.id, 'response', { success: true })
      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'insert_image',
        arguments: {
          source: 'AAAA',
          sourceType: 'base64',
        },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('parseSlideRange', () => {
    it('returns null for undefined input', () => {
      expect(parseSlideRange(undefined)).toBeNull()
    })

    it('returns null for empty string', () => {
      expect(parseSlideRange('')).toBeNull()
    })

    it('parses single index', () => {
      expect(parseSlideRange('5')).toEqual([5])
    })

    it('parses comma-separated indices', () => {
      expect(parseSlideRange('2,4,7')).toEqual([2, 4, 7])
    })

    it('parses a range', () => {
      expect(parseSlideRange('0-3')).toEqual([0, 1, 2, 3])
    })

    it('parses mixed ranges and indices', () => {
      expect(parseSlideRange('0-2,5,8-10')).toEqual([0, 1, 2, 5, 8, 9, 10])
    })

    it('deduplicates overlapping ranges', () => {
      expect(parseSlideRange('0-3,2-5')).toEqual([0, 1, 2, 3, 4, 5])
    })

    it('throws on invalid index', () => {
      expect(() => parseSlideRange('abc')).toThrow('Invalid slide index')
    })

    it('throws on invalid range', () => {
      expect(() => parseSlideRange('5-2')).toThrow('Invalid slide range')
    })

    it('throws on negative index', () => {
      expect(() => parseSlideRange('-1')).toThrow('Invalid slide index')
    })
  })

  describe('preview_deck', () => {
    it('returns interleaved image and text blocks', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'preview_deck',
        arguments: {},
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.action).toBe('executeCode')
      expect(sentJson.params.code).toContain('getImageAsBase64')
      expect(sentJson.params.code).toContain('width: 480')

      pool.handleResponse(sentJson.id, 'response', {
        slideCount: 3,
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            index: 0,
            id: 'slide-0',
            shapeCount: 2,
            shapes: [
              { name: 'Title', type: 'TextBox', id: '1', text: 'Hello World' },
              { name: 'Subtitle', type: 'TextBox', id: '2', text: 'Intro' },
            ],
            imageBase64: 'img0data',
          },
          {
            index: 1,
            id: 'slide-1',
            shapeCount: 1,
            shapes: [{ name: 'Picture', type: 'Image', id: '3' }],
            imageBase64: 'img1data',
          },
          {
            index: 2,
            id: 'slide-2',
            shapeCount: 1,
            shapes: [{ name: 'Body', type: 'TextBox', id: '4', text: 'Content here' }],
            imageBase64: 'img2data',
          },
        ],
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; text?: string; data?: string; mimeType?: string }>

      // Header text
      expect(content[0].type).toBe('text')
      expect(content[0].text).toContain('3 total slides (960 x 540 pt), showing 3')

      // Slide 0: image then text
      expect(content[1].type).toBe('image')
      expect(content[1].data).toBe('img0data')
      expect(content[1].mimeType).toBe('image/png')
      expect(content[2].type).toBe('text')
      expect(content[2].text).toContain('Slide 0')
      expect(content[2].text).toContain('Hello World')
      expect(content[2].text).toContain('Intro')

      // Slide 1: image then text (no text content)
      expect(content[3].type).toBe('image')
      expect(content[4].type).toBe('text')
      expect(content[4].text).toContain('Slide 1')
      expect(content[4].text).toContain('(no text content)')

      // Slide 2: image then text
      expect(content[5].type).toBe('image')
      expect(content[6].type).toBe('text')
      expect(content[6].text).toContain('Content here')
    })

    it('skips images when includeImages is false', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'preview_deck',
        arguments: { includeImages: false },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      // Should NOT contain getImageAsBase64 in the code
      expect(sentJson.params.code).not.toContain('getImageAsBase64')

      pool.handleResponse(sentJson.id, 'response', {
        slideCount: 2,
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            index: 0,
            id: 'slide-0',
            shapeCount: 1,
            shapes: [{ name: 'Title', type: 'TextBox', id: '1', text: 'Slide text' }],
          },
          { index: 1, id: 'slide-1', shapeCount: 0, shapes: [] },
        ],
      })

      const result = await toolPromise
      const content = result.content as Array<{ type: string; text?: string }>

      // Should have no image blocks at all
      const imageBlocks = content.filter((c) => c.type === 'image')
      expect(imageBlocks).toHaveLength(0)

      // Should have header + 2 slide text blocks
      expect(content).toHaveLength(3)
      expect(content[1].text).toContain('Slide text')
    })

    it('passes custom imageWidth to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'preview_deck',
        arguments: { imageWidth: 960 },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('width: 960')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 0, slideWidth: 960, slideHeight: 540, slides: [] })
      await toolPromise
    })

    it('passes slideRange indices to Office.js code', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'preview_deck',
        arguments: { slideRange: '0-2,5' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('[0,1,2,5]')

      pool.handleResponse(sentJson.id, 'response', { slideCount: 10, slideWidth: 960, slideHeight: 540, slides: [] })
      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'preview_deck',
        arguments: {},
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('get_local_copy', () => {
    beforeEach(() => {
      localCopyCache.clear()
    })

    it('returns local file path directly for local files', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: '/path/to/test.pptx',
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'get_local_copy', arguments: {} })
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.localPath).toBe('/path/to/test.pptx')
      expect(parsed.source).toBe('local')
      // No WebSocket commands should have been sent
      expect(ws.send).not.toHaveBeenCalled()
    })

    it('returns error when local file does not exist', async () => {
      const { existsSync } = await import('node:fs')
      vi.mocked(existsSync).mockReturnValueOnce(false)

      const ws = mockWs()
      pool.add('missing.pptx', {
        ws,
        ready: true,
        presentationId: 'missing.pptx',
        filePath: '/path/to/missing.pptx',
      })

      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'get_local_copy', arguments: {} })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('Local file not found')
    })

    it('exports cloud file and writes to temp', async () => {
      const ws = mockWs()
      pool.add('cloud-deck', {
        ws,
        ready: true,
        presentationId: 'cloud-deck',
        filePath: 'https://sharepoint.com/sites/team/Shared%20Documents/deck.pptx',
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({ name: 'get_local_copy', arguments: {} })

      // First command: get revision number
      await new Promise((r) => setTimeout(r, 10))
      const revJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(revJson.params.code).toContain('revisionNumber')
      pool.handleResponse(revJson.id, 'response', 42)

      // Second command: export via getFileAsync
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(exportJson.params.code).toContain('getFileAsync')
      pool.handleResponse(exportJson.id, 'response', 'UEsDBBQAAAA=')

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.source).toBe('exported')
      expect(parsed.revision).toBe(42)
      expect(parsed.localPath).toContain('pptbridge-')
      expect(parsed.localPath).toContain('deck.pptx')

      // Verify writeFileSync was called
      const { writeFileSync } = await import('node:fs')
      expect(writeFileSync).toHaveBeenCalled()
    })

    it('returns cached path when revision unchanged', async () => {
      const ws = mockWs()
      pool.add('cloud-deck', {
        ws,
        ready: true,
        presentationId: 'cloud-deck',
        filePath: 'https://sharepoint.com/sites/team/deck.pptx',
      })

      // Pre-populate cache
      localCopyCache.set('cloud-deck', { localPath: '/tmp/pptbridge-cached-deck.pptx', revision: 7 })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({ name: 'get_local_copy', arguments: {} })

      // Revision check returns same revision
      await new Promise((r) => setTimeout(r, 10))
      const revJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(revJson.id, 'response', 7)

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.source).toBe('cached')
      expect(parsed.localPath).toBe('/tmp/pptbridge-cached-deck.pptx')
      expect(parsed.revision).toBe(7)

      // Should only have sent one command (revision check), not export
      expect(ws.send).toHaveBeenCalledTimes(1)
    })

    it('re-exports when revision has changed', async () => {
      const ws = mockWs()
      pool.add('cloud-deck', {
        ws,
        ready: true,
        presentationId: 'cloud-deck',
        filePath: 'https://sharepoint.com/sites/team/deck.pptx',
      })

      // Pre-populate cache with old revision
      localCopyCache.set('cloud-deck', { localPath: '/tmp/pptbridge-old.pptx', revision: 5 })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({ name: 'get_local_copy', arguments: {} })

      // Revision check returns NEW revision
      await new Promise((r) => setTimeout(r, 10))
      const revJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(revJson.id, 'response', 6)

      // Should trigger export via getFileAsync
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(exportJson.params.code).toContain('getFileAsync')
      pool.handleResponse(exportJson.id, 'response', 'UEsDBBQAAAA=')

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.source).toBe('exported')
      expect(parsed.revision).toBe(6)

      // Two commands sent: revision check + export
      expect(ws.send).toHaveBeenCalledTimes(2)
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({ name: 'get_local_copy', arguments: {} })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('read_shape_paragraphs', () => {
    it('returns paragraph XML for a shape', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'read_shape_paragraphs',
        arguments: { slideIndex: 0, shapeId: '2' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('exportAsBase64')
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('<a:p')
      expect(text).toContain('Hello')
      expect(text).toContain('World')
      expect(text).toContain('b="1"')
      expect(text).not.toContain('<a:bodyPr')
    })

    it('returns error for non-existent shape', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'read_shape_paragraphs',
        arguments: { slideIndex: 0, shapeId: '999' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('not found')
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'read_shape_paragraphs',
        arguments: { slideIndex: 0, shapeId: '2' },
      })
      expect(result.isError).toBe(true)
    })
  })

  describe('edit_shape_paragraphs', () => {
    it('exports, modifies paragraphs, and reimports', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const newParagraphXml =
        '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Replaced</a:t></a:r></a:p>'

      const toolPromise = client.callTool({
        name: 'edit_shape_paragraphs',
        arguments: { slideIndex: 0, shapeId: '2', xml: newParagraphXml },
      })

      // Export command
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(exportJson.params.code).toContain('exportAsBase64')
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // Reimport command
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(reimportJson.params.code).toContain('insertSlidesFromBase64')
      expect(reimportJson.params.code).toContain('slide-0')
      // Verify atomic reimport: delete + insert batched before sync, with post-verification
      const reimportCode = reimportJson.params.code as string
      const deleteIdx = reimportCode.indexOf('.delete()')
      const insertIdx = reimportCode.indexOf('insertSlidesFromBase64')
      const firstSyncAfterDelete = reimportCode.indexOf('await context.sync()', deleteIdx)
      expect(insertIdx).toBeGreaterThan(deleteIdx)
      expect(insertIdx).toBeLessThan(firstSyncAfterDelete)
      expect(reimportCode).toContain('countBefore')
      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.success).toBe(true)
    })

    it('returns error for non-existent shape', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'edit_shape_paragraphs',
        arguments: { slideIndex: 0, shapeId: '999', xml: '<a:p/>' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('not found')
    })
  })

  describe('read_slide_xml', () => {
    it('returns full slide XML when no shapeId', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'read_slide_xml',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('<p:sld')
      expect(text).toContain('Hello')
      expect(text).toContain('Body text')
    })

    it('returns filtered shape XML when shapeId provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'read_slide_xml',
        arguments: { slideIndex: 0, shapeId: '5' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('Content 2')
      expect(text).toContain('Body text')
      // Should NOT contain the other shape
      expect(text).not.toContain('Title 1')
    })

    it('returns error for non-existent shape', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'read_slide_xml',
        arguments: { slideIndex: 0, shapeId: '999' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      expect(result.isError).toBe(true)
    })
  })

  describe('edit_slide_xml', () => {
    it('replaces full slide XML and reimports', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'edit_slide_xml',
        arguments: { slideIndex: 0, xml: SAMPLE_SLIDE_XML.replace('Hello', 'Modified') },
      })

      // Export
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // Reimport
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(reimportJson.params.code).toContain('insertSlidesFromBase64')
      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(JSON.parse(text).success).toBe(true)
    })

    it('replaces specific shape XML when shapeId provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const newShapeXml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="5" name="Replaced"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>New</a:t></a:r></a:p></p:txBody>
      </p:sp>`

      const toolPromise = client.callTool({
        name: 'edit_slide_xml',
        arguments: { slideIndex: 0, xml: newShapeXml, shapeId: '5' },
      })

      // Export
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // Reimport
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      expect(JSON.parse((result.content as Array<{ text: string }>)[0].text).success).toBe(true)
    })

    it('returns error for non-existent shape', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'edit_slide_xml',
        arguments: { slideIndex: 0, xml: '<p:sp/>', shapeId: '999' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      expect(result.isError).toBe(true)
    })

    it('executes code-based DOM manipulation and reimports', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'edit_slide_xml',
        arguments: {
          slideIndex: 0,
          code: `
            var shape = findShapeById("2");
            var spPr = shape.getElementsByTagNameNS(NS_P, "spPr")[0];
            var fill = doc.createElementNS(NS_A, "a:solidFill");
            var clr = doc.createElementNS(NS_A, "a:srgbClr");
            clr.setAttribute("val", "FF0000");
            fill.appendChild(clr);
            spPr.appendChild(fill);
          `,
          explanation: 'Add red fill to title shape',
        },
      })

      // Export
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // Reimport
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(reimportJson.params.code).toContain('insertSlidesFromBase64')
      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      expect(JSON.parse((result.content as Array<{ text: string }>)[0].text).success).toBe(true)
    })

    it('rejects when both xml and code are provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const result = await client.callTool({
        name: 'edit_slide_xml',
        arguments: { slideIndex: 0, xml: '<p:sld/>', code: 'var x = 1;' },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain("Provide either 'xml' or 'code'")
    })

    it('rejects when neither xml nor code is provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const result = await client.callTool({
        name: 'edit_slide_xml',
        arguments: { slideIndex: 0 },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain("Provide either 'xml' or 'code'")
    })

    it('returns error when code throws at runtime', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)
      const base64 = await makeSlideZipBase64()

      const toolPromise = client.callTool({
        name: 'edit_slide_xml',
        arguments: {
          slideIndex: 0,
          code: 'throw new Error("shape not found");',
        },
      })

      // Export
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('Code execution error')
      expect(text).toContain('shape not found')
    })
  })

  describe('format_shapes', () => {
    it('generates Office.js code with setSolidColor for fill', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'format_shapes',
        arguments: {
          slideIndex: 0,
          shapes: [{ id: '2', fill: '1A1A1E' }],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('setSolidColor')
      expect(sentJson.params.code).toContain('1A1A1E')
      pool.handleResponse(sentJson.id, 'response', { success: true })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(JSON.parse(text).success).toBe(true)
    })

    it('generates font property assignments', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'format_shapes',
        arguments: {
          slideIndex: 0,
          shapes: [{ id: '5', font: { bold: true, size: 16, color: 'FFFFFF', name: 'Arial' } }],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      const code = sentJson.params.code
      expect(code).toContain('font.bold = true')
      expect(code).toContain('font.size = 16')
      expect(code).toContain('font.color = "FFFFFF"')
      expect(code).toContain('font.name = "Arial"')
      pool.handleResponse(sentJson.id, 'response', { success: true })

      await toolPromise
    })

    it('generates code for multiple shapes', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'format_shapes',
        arguments: {
          slideIndex: 1,
          shapes: [
            { id: '2', fill: 'FF0000' },
            { id: '5', fill: '00FF00', font: { bold: false } },
          ],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      const code = sentJson.params.code
      expect(code).toContain('"2"')
      expect(code).toContain('"5"')
      expect(code).toContain('slides.items[1]')
      pool.handleResponse(sentJson.id, 'response', { success: true })

      await toolPromise
    })

    it('skips fill code when only font specified', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'format_shapes',
        arguments: {
          slideIndex: 0,
          shapes: [{ id: '2', font: { italic: true } }],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      const code = sentJson.params.code
      expect(code).not.toContain('setSolidColor')
      expect(code).toContain('font.italic = true')
      pool.handleResponse(sentJson.id, 'response', { success: true })

      await toolPromise
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'format_shapes',
        arguments: { slideIndex: 0, shapes: [{ id: '2', fill: 'FF0000' }] },
      })
      expect(result.isError).toBe(true)
    })
  })

  describe('execute_officejs', () => {
    it('sends code through pool and returns result', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      // Start the tool call
      const toolPromise = client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'return 42' },
      })

      // Wait a tick for the command to be sent
      await new Promise((r) => setTimeout(r, 10))

      // Extract and respond to the command
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', 42)

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toBe('42')
    })

    it('returns error on timeout', async () => {
      vi.useFakeTimers()
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'slow code' },
      })

      // Advance past timeout
      await vi.advanceTimersByTimeAsync(200)

      const result = await toolPromise
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('timed out')
      vi.useRealTimers()
    })

    it('returns error when no connections', async () => {
      const { client } = await setupMcpClient(pool)
      const result = await client.callTool({
        name: 'execute_officejs',
        arguments: { code: 'return 1' },
      })
      expect(result.isError).toBe(true)
      const text = (result.content as Array<{ text: string }>)[0].text
      expect(text).toContain('No presentations connected')
    })
  })

  describe('duplicate_slide', () => {
    it('exports and reimports slide at same position', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'duplicate_slide',
        arguments: { slideIndex: 1 },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sentJson.params.code).toContain('exportAsBase64')
      expect(sentJson.params.code).toContain('insertSlidesFromBase64')

      pool.handleResponse(sentJson.id, 'response', {
        duplicatedSlideIndex: 1,
        insertedAfter: 1,
        slideCount: 4,
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.duplicatedSlideIndex).toBe(1)
      expect(parsed.slideCount).toBe(4)
    })

    it('inserts after specified index', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'duplicate_slide',
        arguments: { slideIndex: 0, insertAfter: 3 },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      // The code should reference both index 0 (source) and index 3 (insert position)
      expect(sentJson.params.code).toContain('items[0]')
      expect(sentJson.params.code).toContain('items[3]')

      pool.handleResponse(sentJson.id, 'response', {
        duplicatedSlideIndex: 0,
        insertedAfter: 3,
        slideCount: 5,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.insertedAfter).toBe(3)
    })
  })

  describe('scan_slide', () => {
    it('returns shape objects with correct fields', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'scan_slide',
        arguments: { slideRange: '0' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      pool.handleResponse(sentJson.id, 'response', {
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            slideIndex: 0,
            slideId: 'slide-1',
            shapes: [
              { id: '10', name: 'Title 1', type: 'Rectangle', left: 50, top: 20, width: 400, height: 60 },
              { id: '11', name: 'Content 2', type: 'TextBox', left: 50, top: 100, width: 400, height: 300 },
            ],
          },
        ],
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.slides[0].shapes).toHaveLength(2)
      expect(parsed.slides[0].shapes[0]).toEqual({
        id: '10',
        name: 'Title 1',
        type: 'Rectangle',
        left: 50,
        top: 20,
        width: 400,
        height: 60,
      })
    })

    it('does NOT return text or fill fields', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'scan_slide',
        arguments: { slideRange: '0' },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      // Verify the Office.js code does NOT load text or fill
      expect(sentJson.params.code).not.toContain('textFrame')
      expect(sentJson.params.code).not.toContain('fill')
      expect(sentJson.params.code).not.toContain('textRange')

      pool.handleResponse(sentJson.id, 'response', {
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            slideIndex: 0,
            slideId: 'slide-1',
            shapes: [{ id: '10', name: 'Title 1', type: 'Rectangle', left: 50, top: 20, width: 400, height: 60 }],
          },
        ],
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.slides[0].shapes[0]).not.toHaveProperty('text')
      expect(parsed.slides[0].shapes[0]).not.toHaveProperty('fill')
    })
  })

  describe('verify_slides', () => {
    it('detects overlapping shapes', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'verify_slides',
        arguments: { slideIndex: 0, checks: ['overlap'] },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      pool.handleResponse(sentJson.id, 'response', {
        shapes: [
          { name: 'Shape A', id: '1', left: 0, top: 0, width: 200, height: 100 },
          { name: 'Shape B', id: '2', left: 100, top: 50, width: 200, height: 100 },
          { name: 'Shape C', id: '3', left: 500, top: 500, width: 50, height: 50 },
        ],
        slideWidth: 960,
        slideHeight: 540,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.issueCount).toBe(1)
      expect(parsed.issues[0].type).toBe('overlap')
      expect(parsed.issues[0].shapeIds).toContain('1')
      expect(parsed.issues[0].shapeIds).toContain('2')
    })

    it('detects out-of-bounds shapes', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'verify_slides',
        arguments: { slideIndex: 0, checks: ['bounds'] },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      pool.handleResponse(sentJson.id, 'response', {
        shapes: [
          { name: 'Offscreen', id: '1', left: 900, top: 0, width: 200, height: 100 },
          { name: 'OnScreen', id: '2', left: 100, top: 100, width: 100, height: 100 },
        ],
        slideWidth: 960,
        slideHeight: 540,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.issueCount).toBe(1)
      expect(parsed.issues[0].type).toBe('bounds')
      expect(parsed.issues[0].description).toContain('right of slide')
      expect(parsed.issues[0].shapeIds).toContain('1')
    })

    it('detects empty text and tiny shapes', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'verify_slides',
        arguments: { slideIndex: 0, checks: ['empty_text', 'tiny_shapes'] },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      pool.handleResponse(sentJson.id, 'response', {
        shapes: [
          { name: 'Empty', id: '1', left: 0, top: 0, width: 100, height: 50, text: '  ' },
          { name: 'Tiny', id: '2', left: 0, top: 0, width: 5, height: 3 },
          { name: 'Normal', id: '3', left: 0, top: 0, width: 200, height: 100, text: 'Hello' },
        ],
        slideWidth: 960,
        slideHeight: 540,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.issueCount).toBe(2)
      const types = parsed.issues.map((i: { type: string }) => i.type)
      expect(types).toContain('empty_text')
      expect(types).toContain('tiny_shapes')
      // Verify shapeIds contains numeric IDs, not names
      const emptyTextIssue = parsed.issues.find((i: { type: string }) => i.type === 'empty_text')
      expect(emptyTextIssue.shapeIds).toContain('1')
      const tinyIssue = parsed.issues.find((i: { type: string }) => i.type === 'tiny_shapes')
      expect(tinyIssue.shapeIds).toContain('2')
    })

    it('runs all checks by default', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const toolPromise = client.callTool({
        name: 'verify_slides',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])

      pool.handleResponse(sentJson.id, 'response', {
        shapes: [{ name: 'Good', id: '1', left: 100, top: 100, width: 200, height: 100, text: 'OK' }],
        slideWidth: 960,
        slideHeight: 540,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.issueCount).toBe(0)
      expect(parsed.shapeCount).toBe(1)
    })
  })

  describe('read_slide_zip', () => {
    it('returns discovered files when no paths specified', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types/>')
      zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
      zip.file('ppt/media/image1.png', 'binarydata')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'read_slide_zip',
        arguments: { slideIndex: 0 },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', {
        base64,
        slideId: 'slide-0',
        prevSlideId: null,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      // Should include xml/rels files but not binary
      expect(parsed.zipContents['ppt/slides/slide1.xml']).toContain('Hello')
      expect(parsed.zipContents['[Content_Types].xml']).toBeDefined()
      expect(parsed.zipContents['ppt/slides/_rels/slide1.xml.rels']).toBeDefined()
      expect(parsed.zipContents['ppt/media/image1.png']).toBeUndefined()
      // allPaths includes everything
      expect(parsed.allPaths).toContain('ppt/media/image1.png')
    })

    it('returns specific files when paths provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('ppt/charts/chart1.xml', '<c:chartSpace/>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'read_slide_zip',
        arguments: { slideIndex: 0, paths: ['ppt/charts/chart1.xml'] },
      })

      await new Promise((r) => setTimeout(r, 10))
      const sentJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(sentJson.id, 'response', {
        base64,
        slideId: 'slide-0',
        prevSlideId: null,
      })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(Object.keys(parsed.zipContents)).toHaveLength(1)
      expect(parsed.zipContents['ppt/charts/chart1.xml']).toContain('chartSpace')
    })
  })

  describe('edit_slide_zip', () => {
    it('updates files and reimports', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types></Types>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'edit_slide_zip',
        arguments: {
          slideIndex: 0,
          files: { 'ppt/slides/slide1.xml': SAMPLE_SLIDE_XML.replace('Hello', 'Updated') },
        },
      })

      // First call: exportSlide
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(exportJson.params.code).toContain('exportAsBase64')
      pool.handleResponse(exportJson.id, 'response', {
        base64,
        slideId: 'slide-0',
        prevSlideId: null,
      })

      // Second call: reimportSlide
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      expect(reimportJson.params.code).toContain('insertSlidesFromBase64')
      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.success).toBe(true)
      expect(parsed.filesUpdated).toBe(1)
      expect(parsed.newFiles).toHaveLength(0)
    })

    it('auto-registers Content_Types for new chart files', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'edit_slide_zip',
        arguments: {
          slideIndex: 0,
          files: {
            'ppt/slides/slide1.xml': SAMPLE_SLIDE_XML,
            'ppt/charts/chart1.xml': '<c:chartSpace/>',
          },
        },
      })

      // exportSlide
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // reimportSlide — the reimported base64 should contain the chart + updated Content_Types
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      // Verify the reimported base64 contains the auto-registered Content_Types
      const reimportCode = reimportJson.params.code as string
      // Extract the base64 from insertSlidesFromBase64("...") call
      const b64Match = reimportCode.match(/insertSlidesFromBase64\("([^"]+)"/)
      expect(b64Match).not.toBeNull()
      const reimportedZip = new JSZip()
      await reimportedZip.loadAsync(b64Match![1], { base64: true })
      const ct = await reimportedZip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('PartName="/ppt/charts/chart1.xml"')
      expect(ct).toContain('drawingml.chart+xml')
      // Verify chart file was added
      const chart = await reimportedZip.file('ppt/charts/chart1.xml')!.async('string')
      expect(chart).toContain('chartSpace')

      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.success).toBe(true)
      expect(parsed.newFiles).toContain('ppt/charts/chart1.xml')
    })

    it('skips auto-registration when Content_Types is explicitly provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types></Types>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const customCt = '<Types><Override PartName="/ppt/charts/chart1.xml" ContentType="custom/type"/></Types>'

      const toolPromise = client.callTool({
        name: 'edit_slide_zip',
        arguments: {
          slideIndex: 0,
          files: {
            'ppt/charts/chart1.xml': '<c:chartSpace/>',
            '[Content_Types].xml': customCt,
          },
        },
      })

      // exportSlide
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // reimportSlide — should use the explicit Content_Types, not auto-registered
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      const reimportCode = reimportJson.params.code as string
      const b64Match = reimportCode.match(/insertSlidesFromBase64\("([^"]+)"/)
      const reimportedZip = new JSZip()
      await reimportedZip.loadAsync(b64Match![1], { base64: true })
      const ct = await reimportedZip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('custom/type')

      pool.handleResponse(reimportJson.id, 'response', { success: true })
      await toolPromise
    })
  })

  describe('edit_slide_chart', () => {
    it('creates a column chart with all required OOXML parts', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      // Prepare a zip with slide XML, rels, and Content_Types
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file(
        'ppt/slides/_rels/slide1.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
      )
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'edit_slide_chart',
        arguments: {
          slideIndex: 0,
          chartType: 'column',
          title: 'Revenue by Quarter',
          categories: ['Q1', 'Q2', 'Q3'],
          series: [{ name: 'Revenue', values: [100, 150, 120] }],
        },
      })

      // exportSlide
      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      // reimportSlide — verify the reimported zip has chart, rels, graphic frame
      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      const reimportCode = reimportJson.params.code as string
      const b64Match = reimportCode.match(/insertSlidesFromBase64\("([^"]+)"/)
      expect(b64Match).not.toBeNull()

      const reimportedZip = new JSZip()
      await reimportedZip.loadAsync(b64Match![1], { base64: true })

      // Verify chart XML was created
      const chartFile = reimportedZip.file('ppt/charts/chart1.xml')
      expect(chartFile).not.toBeNull()
      const chartXml = await chartFile!.async('string')
      expect(chartXml).toContain('<c:chartSpace')
      expect(chartXml).toContain('<c:style val="2"/>')
      expect(chartXml).toContain('<c:barChart>')
      expect(chartXml).toContain('<c:barDir val="col"/>')
      expect(chartXml).toContain('Revenue by Quarter')
      expect(chartXml).toContain('<c:dLbls>')

      // Verify rels was updated with chart relationship
      const rels = await reimportedZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
      expect(rels).toContain('rId2')
      expect(rels).toContain('../charts/chart1.xml')
      expect(rels).toContain('relationships/chart')

      // Verify slide XML has graphic frame
      const slideXml = await reimportedZip.file('ppt/slides/slide1.xml')!.async('string')
      expect(slideXml).toContain('<p:graphicFrame')
      expect(slideXml).toContain('r:id="rId2"')

      // Verify Content_Types was auto-registered
      const ct = await reimportedZip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('chart1.xml')
      expect(ct).toContain('drawingml.chart+xml')

      pool.handleResponse(reimportJson.id, 'response', { success: true })

      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.success).toBe(true)
      expect(parsed.chartType).toBe('column')
      expect(parsed.chartFile).toBe('ppt/charts/chart1.xml')
    })

    it('creates a pie chart without axes', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file(
        'ppt/slides/_rels/slide1.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
      )
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'edit_slide_chart',
        arguments: {
          slideIndex: 0,
          chartType: 'pie',
          title: 'Market Share',
          categories: ['A', 'B', 'C'],
          series: [{ name: 'Share', values: [40, 35, 25] }],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      const reimportCode = reimportJson.params.code as string
      const b64Match = reimportCode.match(/insertSlidesFromBase64\("([^"]+)"/)

      const reimportedZip = new JSZip()
      await reimportedZip.loadAsync(b64Match![1], { base64: true })
      const chartXml = await reimportedZip.file('ppt/charts/chart1.xml')!.async('string')
      expect(chartXml).toContain('<c:pieChart>')
      expect(chartXml).not.toContain('<c:catAx>')
      expect(chartXml).toContain('<c:showPercent val="1"/>')

      pool.handleResponse(reimportJson.id, 'response', { success: true })
      await toolPromise
    })

    it('increments chart number when charts already exist', async () => {
      const ws = mockWs()
      pool.add('test.pptx', { ws, ready: true, presentationId: 'test.pptx', filePath: null })
      const { client } = await setupMcpClient(pool)

      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('ppt/charts/chart1.xml', '<c:chartSpace/>') // existing chart
      zip.file(
        'ppt/slides/_rels/slide1.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="y"/></Relationships>',
      )
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      const base64 = await zip.generateAsync({ type: 'base64' })

      const toolPromise = client.callTool({
        name: 'edit_slide_chart',
        arguments: {
          slideIndex: 0,
          chartType: 'line',
          title: 'Trends',
          categories: ['Jan', 'Feb'],
          series: [{ name: 'Sales', values: [10, 20] }],
        },
      })

      await new Promise((r) => setTimeout(r, 10))
      const exportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      pool.handleResponse(exportJson.id, 'response', { base64, slideId: 'slide-0', prevSlideId: null })

      await new Promise((r) => setTimeout(r, 10))
      const reimportJson = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[1][0])
      const reimportCode = reimportJson.params.code as string
      const b64Match = reimportCode.match(/insertSlidesFromBase64\("([^"]+)"/)

      const reimportedZip = new JSZip()
      await reimportedZip.loadAsync(b64Match![1], { base64: true })

      // chart2.xml should be created (chart1 already exists)
      expect(reimportedZip.file('ppt/charts/chart2.xml')).not.toBeNull()
      // Rels should use rId2 (rId1 already exists)
      const rels = await reimportedZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
      expect(rels).toContain('rId2')
      expect(rels).toContain('../charts/chart2.xml')

      pool.handleResponse(reimportJson.id, 'response', { success: true })
      const result = await toolPromise
      const parsed = JSON.parse((result.content as Array<{ text: string }>)[0].text)
      expect(parsed.chartFile).toBe('ppt/charts/chart2.xml')
    })
  })

  describe('search_text', () => {
    it('searches all slides and returns shape-level matches by default', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: '/path/test.pptx',
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'budget' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.action).toBe('executeCode')
      expect(sent.params.code).toContain('budget')

      pool.handleResponse(sent.id, 'response', {
        query: 'budget',
        caseSensitive: false,
        regex: false,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 1,
            shapeId: '42',
            shapeName: 'TextBox 3',
            source: 'shape',
            text: 'The budget for Q3 is $1.2M',
          },
          {
            slideIndex: 3,
            shapeId: '78',
            shapeName: 'Content Placeholder',
            source: 'shape',
            text: 'Budget allocation overview',
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.query).toBe('budget')
      expect(parsed.matches).toHaveLength(2)
      expect(parsed.matches[0].slideIndex).toBe(1)
      expect(parsed.matches[0].shapeId).toBe('42')
      expect(parsed.matches[0].source).toBe('shape')
      expect(parsed.matches[1].slideIndex).toBe(3)
    })

    it('searches specific slide range when slideRange is provided', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'hello', slideRange: '2-4' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('2-4')

      pool.handleResponse(sent.id, 'response', {
        query: 'hello',
        caseSensitive: false,
        regex: false,
        totalSlides: 10,
        matches: [],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matches).toHaveLength(0)
    })

    it('supports case-sensitive search', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'Budget', caseSensitive: true },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('caseSensitive = true')

      pool.handleResponse(sent.id, 'response', {
        query: 'Budget',
        caseSensitive: true,
        regex: false,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 3,
            shapeId: '78',
            shapeName: 'Title',
            source: 'shape',
            text: 'Budget allocation overview',
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.caseSensitive).toBe(true)
      expect(parsed.matches).toHaveLength(1)
    })

    it('enables regex matching when regex=true', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: '\\d+%', regex: true },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('useRegex = true')
      expect(sent.params.code).toContain('new RegExp')

      pool.handleResponse(sent.id, 'response', {
        query: '\\d+%',
        caseSensitive: false,
        regex: true,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 2,
            shapeId: '10',
            shapeName: 'Body',
            source: 'shape',
            text: 'Revenue grew 42% YoY',
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.regex).toBe(true)
      expect(parsed.matches).toHaveLength(1)
    })

    it('returns slide-level context with all shapes when context="slide"', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'budget', context: 'slide' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('"slide"')

      pool.handleResponse(sent.id, 'response', {
        query: 'budget',
        caseSensitive: false,
        regex: false,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 1,
            shapes: [
              { shapeId: '2', shapeName: 'Title', matched: false, text: 'Q3 Overview' },
              { shapeId: '4', shapeName: 'Body', matched: true, text: 'The budget is $1.2M' },
              { shapeId: '7', shapeName: 'Footer', matched: false, text: 'Page 2' },
            ],
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matches).toHaveLength(1)
      expect(parsed.matches[0].slideIndex).toBe(1)
      expect(parsed.matches[0].shapes).toHaveLength(3)
      expect(parsed.matches[0].shapes[0].matched).toBe(false)
      expect(parsed.matches[0].shapes[1].matched).toBe(true)
    })

    it('returns only slide indices when context="none"', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'AI', context: 'none' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('"none"')

      pool.handleResponse(sent.id, 'response', {
        query: 'AI',
        caseSensitive: false,
        regex: false,
        totalSlides: 10,
        matchingSlides: [0, 1, 3, 5, 7],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matchingSlides).toEqual([0, 1, 3, 5, 7])
      expect(parsed.matches).toBeUndefined()
    })

    it('includes speaker notes matches when includeNotes is true', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'reminder' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('searchNotes = true')
      expect(sent.params.code).toContain('notesSlide')

      pool.handleResponse(sent.id, 'response', {
        query: 'reminder',
        caseSensitive: false,
        regex: false,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 2,
            source: 'note',
            text: 'Reminder: mention the Q3 deadline',
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matches).toHaveLength(1)
      expect(parsed.matches[0].source).toBe('note')
      expect(parsed.matches[0].slideIndex).toBe(2)
    })

    it('skips notes search when includeNotes is false', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'test', includeNotes: false },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('searchNotes = false')

      pool.handleResponse(sent.id, 'response', {
        query: 'test',
        caseSensitive: false,
        regex: false,
        totalSlides: 5,
        matches: [],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matches).toHaveLength(0)
    })

    it('includes table cell matches with row/col info', async () => {
      const ws = mockWs()
      pool.add('test.pptx', {
        ws,
        ready: true,
        presentationId: 'test.pptx',
        filePath: null,
      })

      const { client } = await setupMcpClient(pool)
      const toolPromise = client.callTool({
        name: 'search_text',
        arguments: { query: 'revenue' },
      })

      await new Promise((r) => setTimeout(r, 10))

      const sent = JSON.parse((ws.send as ReturnType<typeof vi.fn>).mock.calls[0][0])
      expect(sent.params.code).toContain('Table')
      expect(sent.params.code).toContain('getCell')

      pool.handleResponse(sent.id, 'response', {
        query: 'revenue',
        caseSensitive: false,
        regex: false,
        totalSlides: 5,
        matches: [
          {
            slideIndex: 4,
            shapeId: '20',
            shapeName: 'Table 1',
            source: 'tableCell',
            text: 'Total revenue: $5M',
            row: 2,
            col: 1,
          },
        ],
      })

      const result = await toolPromise
      const text = (result.content as Array<{ text: string }>)[0].text
      const parsed = JSON.parse(text)
      expect(parsed.matches).toHaveLength(1)
      expect(parsed.matches[0].source).toBe('tableCell')
      expect(parsed.matches[0].row).toBe(2)
      expect(parsed.matches[0].col).toBe(1)
    })
  })
})
