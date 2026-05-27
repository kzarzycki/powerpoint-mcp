import { readFileSync } from 'node:fs'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { recolorSvg, searchIcons } from '../icons.ts'
import { withTool } from '../tool-helpers.ts'
import { getConcurrentWarning, getLocalCopyPath, localCopyCache } from './shared.ts'

export function registerMediaTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'insert_image',
    'Inserts an image onto a slide using Office.js setSelectedDataAsync with CoercionType.Image. Accepts a file path, URL, or raw base64 data. Optionally navigate to a specific slide first and control position/size in points.',
    {
      source: z.string().describe('File path, URL, or base64 image data depending on sourceType'),
      sourceType: z
        .enum(['file', 'url', 'base64'])
        .describe(
          'How to interpret source: "file" reads from disk, "url" fetches from network, "base64" uses data directly',
        ),
      slideIndex: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          'Zero-based slide index to navigate to before inserting. If omitted, inserts on the currently active slide.',
        ),
      left: z.number().optional().describe('Horizontal position in points (1 point = 1/72 inch)'),
      top: z.number().optional().describe('Vertical position in points'),
      width: z.number().optional().describe('Image width in points'),
      height: z.number().optional().describe('Image height in points'),
      color: z
        .string()
        .optional()
        .describe(
          'Hex color to tint SVG images (e.g. "#FF5733"). Only applies to SVG sources. Works best with mono/outline icons.',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ source, sourceType, slideIndex, left, top, width, height, color, presentationId }) => {
      // Step 1: Resolve image to base64
      let base64Data: string
      if (sourceType === 'file') {
        base64Data = readFileSync(source).toString('base64')
      } else if (sourceType === 'url') {
        const resp = await fetch(source)
        if (!resp.ok) {
          throw new Error(`Failed to fetch image from URL: ${resp.status} ${resp.statusText}`)
        }
        const buf = await resp.arrayBuffer()
        base64Data = Buffer.from(buf).toString('base64')
      } else {
        base64Data = source
      }

      // Step 1b: Recolor SVG if color is provided
      if (color) {
        const svg = Buffer.from(base64Data, 'base64').toString('utf-8')
        if (svg.trimStart().startsWith('<svg') || svg.trimStart().startsWith('<?xml')) {
          base64Data = Buffer.from(recolorSvg(svg, color)).toString('base64')
        } else {
          throw new Error('color parameter only works with SVG images, but the source is not SVG')
        }
      }

      // Step 2: Build options object string with only provided params
      const optionsParts: string[] = ['coercionType: Office.CoercionType.Image']
      if (left !== undefined) optionsParts.push(`imageLeft: ${left}`)
      if (top !== undefined) optionsParts.push(`imageTop: ${top}`)
      if (width !== undefined) optionsParts.push(`imageWidth: ${width}`)
      if (height !== undefined) optionsParts.push(`imageHeight: ${height}`)
      const optionsStr = `{ ${optionsParts.join(', ')} }`

      // Step 3: Build the setSelectedDataAsync call
      const insertCall = `Office.context.document.setSelectedDataAsync("${base64Data}", ${optionsStr}, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ success: true });
        } else {
          reject(new Error(result.error.message));
        }
      });`

      // Step 4: Wrap with goToByIdAsync if slideIndex is provided
      let code: string
      if (slideIndex !== undefined) {
        code = `return new Promise(function(resolve, reject) {
      Office.context.document.goToByIdAsync(${slideIndex + 1}, Office.GoToType.Index, function(navResult) {
        if (navResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Navigation failed: " + navResult.error.message));
          return;
        }
        ${insertCall}
      });
    });`
      } else {
        code = `return new Promise(function(resolve, reject) {
      ${insertCall}
    });`
      }

      const target = pool.resolveTarget(presentationId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'get_local_copy',
    'Returns a local file path for the presentation. For local files, returns the existing path. For SharePoint/cloud files, exports server-side and saves to a temp .pptx. Caches by revision number — re-exports only when the presentation has been saved since last export.',
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ presentationId }) => {
      const target = pool.resolveTarget(presentationId)
      const cachedBefore = localCopyCache.get(target.presentationId)?.localPath
      const localPath = await getLocalCopyPath(pool, target)
      const isLocal = target.filePath && !target.filePath.startsWith('http')
      const cached = localCopyCache.get(target.presentationId)
      const source = isLocal ? 'local' : cachedBefore === localPath ? 'cached' : 'exported'
      const result: Record<string, unknown> = { localPath, source }
      if (cached) result.revision = cached.revision
      return { content: [{ type: 'text' as const, text: JSON.stringify(result) }] }
    }),
  )

  server.tool(
    'search_fluent_icons',
    'Search Microsoft Fluent UI icon library. Returns matching icons with SVG URLs for use with insert_image. Prefer mono (_M) variants for professional decks. Retry with synonyms if no good matches (e.g. "innovation" → "lightbulb", "security" → "shield").',
    {
      query: z.string().describe('Search query (e.g. "warning", "arrow down", "lightbulb")'),
      top: z.number().int().min(1).max(50).optional().describe('Max results to return (default 10)'),
      style: z
        .enum(['regular', 'filled'])
        .optional()
        .describe('Filter by style: "regular" for mono/outline icons, "filled" for solid icons. Omit for both.'),
    },
    withTool(async ({ query, top, style }) => {
      const results = await searchIcons(query, top ?? 10, style)
      return { content: [{ type: 'text' as const, text: JSON.stringify(results, null, 2) }] }
    }),
  )
}
