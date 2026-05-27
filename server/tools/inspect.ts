import { readFileSync } from 'node:fs'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import JSZip from 'jszip'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import { exportSlide, extractLayoutsFromZip, extractThemeFromZip } from '../xml-helpers.ts'
import {
  getConcurrentWarning,
  getLayoutUsage,
  getLocalCopyPath,
  globToRegExp,
  parseSlideRange,
  themeCache,
} from './shared.ts'

export function registerInspectTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'list_presentations',
    'Lists all PowerPoint presentations currently connected to the bridge server. Shows presentation IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the presentationId to pass to other tools when multiple presentations are open.',
    withTool(async () => {
      const presentations = []
      for (const [id, conn] of pool.entries()) {
        presentations.push({
          presentationId: id,
          filePath: conn.filePath,
          ready: conn.ready,
        })
      }
      return {
        content: [
          {
            type: 'text' as const,
            text:
              presentations.length === 0
                ? 'No presentations connected. Open a PowerPoint file with the bridge add-in loaded.'
                : JSON.stringify(presentations, null, 2),
          },
        ],
      }
    }),
  )

  server.tool(
    'inspect_deck',
    'Deck overview: slide dimensions, theme (colors + fonts), and all slides with index, ID, and shape count. Use as the first call to understand deck structure. Theme is cached after the first call. For shape details, follow up with scan_slide or inspect_slide on specific slides. For available layouts, use inspect_layouts.',
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ presentationId }) => {
      const code = `
          var p = context.presentation;
          var slides = p.slides;
          var ps = p.pageSetup;
          slides.load("items");
          ps.load("slideWidth,slideHeight");
          await context.sync();
          var output = [];
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            slide.shapes.load("items");
            slide.layout.load("name");
          }
          await context.sync();
          for (var i = 0; i < slides.items.length; i++) {
            var slide = slides.items[i];
            output.push({ index: i, id: slide.id, layout: slide.layout.name, shapeCount: slide.shapes.items.length });
          }
          return { slideWidth: ps.slideWidth, slideHeight: ps.slideHeight, slides: output };
        `
      const target = pool.resolveTarget(presentationId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)

      // Extract theme (cached per presentation)
      let theme = themeCache.get(target.presentationId)
      if (!theme) {
        try {
          const exported = await exportSlide(pool, 0, target.ws)
          theme = await extractThemeFromZip(exported.base64)
          themeCache.set(target.presentationId, theme)
        } catch {
          // Theme extraction is best-effort — don't fail the whole call
        }
      }

      const output = { ...(result as Record<string, unknown>), ...(theme ? { theme } : {}) }
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(output) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'inspect_layouts',
    'Returns slide layouts with names, OOXML type (e.g. blank, twoObj, secHead), indices (for slides.add({ layoutIndex })), and detailed placeholders. Use `fields` to control which data is returned. By default reads all layouts from OOXML (complete list, requires file access — may take a moment on first call for cloud files). Set usedOnly to return only layouts assigned to existing slides (fast, Office.js only, no file access).',
    {
      fields: z
        .string()
        .optional()
        .describe(
          'Comma-separated layout fields to include. Placeholders sub-fields in parens. Default: "index,name,type,usedBySlides,placeholders(type,idx,name)". All placeholder fields: type,idx,name,description,sz,left,top,width,height.',
        ),
      usedOnly: z
        .boolean()
        .optional()
        .describe(
          'If true, return only layouts currently assigned to slides (fast, Office.js only). Default: false (all layouts from OOXML).',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ fields, usedOnly, presentationId }) => {
      // Parse fields spec: "index,name,placeholders(type,idx,name)" → { layout: Set, placeholder: Set }
      const DEFAULT_FIELDS = 'index,name,type,usedBySlides,placeholders(type,idx,name)'
      const fieldSpec = fields ?? DEFAULT_FIELDS
      const phMatch = fieldSpec.match(/placeholders\(([^)]+)\)/)
      const phFields = phMatch ? new Set(phMatch[1]!.split(',').map((f) => f.trim())) : null
      const layoutFields = new Set(
        fieldSpec
          .replace(/placeholders\([^)]*\)/, 'placeholders')
          .split(',')
          .map((f) => f.trim()),
      )

      function filterLayout(layout: Record<string, unknown>): Record<string, unknown> {
        const out: Record<string, unknown> = {}
        for (const key of layoutFields) {
          if (key === 'placeholders' && phFields && Array.isArray(layout.placeholders)) {
            out.placeholders = (layout.placeholders as Record<string, unknown>[]).map((ph) => {
              const filtered: Record<string, unknown> = {}
              for (const f of phFields) {
                if (ph[f] !== undefined) filtered[f] = ph[f]
              }
              return filtered
            })
          } else if (key === 'usedBySlides' && Array.isArray(layout[key]) && (layout[key] as unknown[]).length === 0) {
            // Omit empty usedBySlides to reduce payload
          } else if (layout[key] !== undefined) {
            out[key] = layout[key]
          }
        }
        return out
      }

      const target = pool.resolveTarget(presentationId)

      if (usedOnly) {
        // Fast path: just Office.js enumeration
        const layouts = await getLayoutUsage(pool, target.ws)
        return { content: [{ type: 'text' as const, text: JSON.stringify({ layouts, usedOnly: true }) }] }
      }

      // Full path: read all layouts from OOXML via local file
      const localPath = await getLocalCopyPath(pool, target)
      const fileData = readFileSync(localPath)
      const zip = await JSZip.loadAsync(fileData)
      const layouts = await extractLayoutsFromZip(zip)

      // Enrich with usedBySlides from Office.js (best-effort)
      try {
        const usage = await getLayoutUsage(pool, target.ws)
        const usageByName = new Map(usage.map((u) => [u.name, u.usedBySlides]))
        for (const layout of layouts) {
          layout.usedBySlides = usageByName.get(layout.name) ?? []
        }
      } catch {
        // If Office.js enumeration fails, return layouts without usage info
      }

      const filtered = layouts.map((l) => filterLayout(l as unknown as Record<string, unknown>))
      return { content: [{ type: 'text' as const, text: JSON.stringify({ layouts: filtered }) }] }
    }),
  )

  server.tool(
    'inspect_slide',
    'Detailed slide inspector (~80 tokens/shape): returns every shape with text content, positions, sizes, and fill colors, plus slide dimensions. Supports slideRange for multiple slides. For just positions without text/fills, use scan_slide instead.',
    {
      slideRange: z.string().describe('Slide indices to inspect, e.g. "0", "0-5", "2,4,7". Single index or range.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideRange, presentationId }) => {
      const indices = parseSlideRange(slideRange) ?? []
      if (indices.length === 0) throw new Error('slideRange is required')
      const indicesJs = JSON.stringify(indices)
      const code = `
          var p = context.presentation;
          var slides = p.slides;
          var ps = p.pageSetup;
          slides.load("items");
          ps.load("slideWidth,slideHeight");
          await context.sync();
          var requestedIndices = ${indicesJs};
          for (var i = 0; i < requestedIndices.length; i++) {
            if (requestedIndices[i] >= slides.items.length) {
              throw new Error("Slide index " + requestedIndices[i] + " out of range (presentation has " + slides.items.length + " slides)");
            }
          }
          for (var i = 0; i < requestedIndices.length; i++) {
            slides.items[requestedIndices[i]].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < requestedIndices.length; i++) {
            var idx = requestedIndices[i];
            var slide = slides.items[idx];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              var info = {
                name: s.name,
                type: s.type,
                id: s.id,
                left: s.left,
                top: s.top,
                width: s.width,
                height: s.height
              };
              try {
                s.textFrame.load("textRange");
                await context.sync();
                info.text = s.textFrame.textRange.text;
              } catch (e) {
                // Shape has no text frame (e.g., images, connectors)
              }
              try {
                s.fill.load("foregroundColor,type");
                await context.sync();
                info.fill = { type: s.fill.type, color: s.fill.foregroundColor };
              } catch (e) {
                // Shape has no fill or fill not accessible
              }
              shapes.push(info);
            }
            output.push({ slideIndex: idx, slideId: slide.id, shapes: shapes });
          }
          return { slideWidth: ps.slideWidth, slideHeight: ps.slideHeight, slides: output };
        `
      const target = pool.resolveTarget(presentationId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(result, null, 2) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'scan_slide',
    'Lightweight shape scanner (~40 tokens/shape): lists shape IDs, types, and positions on slides, plus slide dimensions. Supports slideRange for multiple slides. For text content and fills, use inspect_slide instead.',
    {
      slideRange: z.string().describe('Slide indices to scan, e.g. "0", "0-5", "2,4,7". Single index or range.'),
      namePattern: z
        .string()
        .optional()
        .describe('Glob-style filter on shape name, e.g. "Title*", "*_source", "Card*_bg". Case-insensitive.'),
      shapeType: z
        .string()
        .optional()
        .describe('Filter by shape type: "Placeholder", "TextBox", "GeometricShape", "Graphic", "Picture", etc.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideRange, namePattern, shapeType, presentationId }) => {
      const indices = parseSlideRange(slideRange) ?? []
      if (indices.length === 0) throw new Error('slideRange is required')
      const indicesJs = JSON.stringify(indices)
      const code = `
          var p = context.presentation;
          var slides = p.slides;
          var ps = p.pageSetup;
          slides.load("items");
          ps.load("slideWidth,slideHeight");
          await context.sync();
          var requestedIndices = ${indicesJs};
          for (var i = 0; i < requestedIndices.length; i++) {
            if (requestedIndices[i] >= slides.items.length) {
              throw new Error("Slide index " + requestedIndices[i] + " out of range (presentation has " + slides.items.length + " slides)");
            }
          }
          for (var i = 0; i < requestedIndices.length; i++) {
            slides.items[requestedIndices[i]].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < requestedIndices.length; i++) {
            var idx = requestedIndices[i];
            var slide = slides.items[idx];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              shapes.push({
                id: s.id,
                name: s.name,
                type: s.type,
                left: s.left,
                top: s.top,
                width: s.width,
                height: s.height
              });
            }
            output.push({ slideIndex: idx, slideId: slide.id, shapes: shapes });
          }
          return { slideWidth: ps.slideWidth, slideHeight: ps.slideHeight, slides: output };
        `
      const target = pool.resolveTarget(presentationId)
      const result = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
        slideWidth: number
        slideHeight: number
        slides: Array<{
          slideIndex: number
          slideId: string
          shapes: Array<{
            id: string
            name: string
            type: string
            left: number
            top: number
            width: number
            height: number
          }>
        }>
      }

      // Apply optional filters (post-processing, no extra Office.js calls)
      if (namePattern || shapeType) {
        const nameRegex = namePattern ? globToRegExp(namePattern) : null
        const typeLower = shapeType?.toLowerCase()
        for (const slide of result.slides) {
          slide.shapes = slide.shapes.filter((s) => {
            if (nameRegex && !nameRegex.test(s.name)) return false
            if (typeLower && s.type.toLowerCase() !== typeLower) return false
            return true
          })
        }
      }

      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(result, null, 2) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'screenshot_slide',
    'Slide screenshot (~1000 tokens): captures one slide as PNG image. Use to visually verify layout after changes. Do NOT loop over all slides — use preview_deck instead.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from scan_slide results'),
      width: z
        .number()
        .int()
        .min(1)
        .max(4096)
        .optional()
        .describe(
          'Image width in pixels. Default: 720. Height auto-calculated to preserve aspect ratio unless also specified.',
        ),
      height: z
        .number()
        .int()
        .min(1)
        .max(4096)
        .optional()
        .describe('Image height in pixels. If omitted, auto-calculated from width to preserve aspect ratio.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideIndex, width, height, presentationId }) => {
      const imgWidth = width ?? 720
      const optionsParts: string[] = [`width: ${imgWidth}`]
      if (height !== undefined) {
        optionsParts.push(`height: ${height}`)
      }
      const optionsStr = `{ ${optionsParts.join(', ')} }`

      const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          var result = slide.getImageAsBase64(${optionsStr});
          await context.sync();
          return { base64: result.value, slideIndex: ${slideIndex}, slideId: slide.id };
        `
      try {
        const target = pool.resolveTarget(presentationId)
        const result = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
          base64: string
          slideIndex: number
          slideId: string
        }
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const description = `Slide ${result.slideIndex} (ID: ${result.slideId})${warning ?? ''}`

        return {
          content: [
            {
              type: 'image' as const,
              data: result.base64,
              mimeType: 'image/png',
            },
            {
              type: 'text' as const,
              text: description,
            },
          ],
        }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        if (message.includes('getImageAsBase64') || message.includes('not a function')) {
          throw new Error(`${message} (This API requires PowerPoint 16.96+ with PowerPointApi 1.8 support)`)
        }
        throw err
      }
    }),
  )

  server.tool(
    'preview_deck',
    'Deck preview: batch overview of all/selected slides with optional thumbnails + text. With images: ~900 tokens/slide; text-only (includeImages=false): ~35 tokens/slide. Use for visual review or content audit. Do NOT use to inspect one slide — use inspect_slide or screenshot_slide instead.',
    {
      slideRange: z
        .string()
        .optional()
        .describe('Slide indices to include, e.g. "0-5", "2,4,7", "0-2,5,8-10". Omit for all slides.'),
      imageWidth: z
        .number()
        .int()
        .min(120)
        .max(1920)
        .optional()
        .describe('Thumbnail width in pixels. Default: 480. Height auto-calculated to preserve aspect ratio.'),
      includeImages: z
        .boolean()
        .optional()
        .describe('Include slide thumbnails. Default: true. Set false for text-only overview (faster).'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideRange, imageWidth, includeImages, presentationId }) => {
      const indices = parseSlideRange(slideRange)
      const width = imageWidth ?? 480
      const withImages = includeImages !== false

      // Build the indices array literal for Office.js, or null for "all"
      const indicesJs = indices ? JSON.stringify(indices) : 'null'

      const code = `
          var p = context.presentation;
          var slides = p.slides;
          var ps = p.pageSetup;
          slides.load("items");
          ps.load("slideWidth,slideHeight");
          await context.sync();
          var requestedIndices = ${indicesJs};
          var indicesToProcess = requestedIndices || [];
          if (!requestedIndices) {
            for (var i = 0; i < slides.items.length; i++) indicesToProcess.push(i);
          }
          // Validate indices
          for (var i = 0; i < indicesToProcess.length; i++) {
            if (indicesToProcess[i] >= slides.items.length) {
              throw new Error("Slide index " + indicesToProcess[i] + " out of range (presentation has " + slides.items.length + " slides)");
            }
          }
          // Load shapes for all requested slides
          for (var i = 0; i < indicesToProcess.length; i++) {
            slides.items[indicesToProcess[i]].shapes.load("items");
          }
          await context.sync();
          var output = [];
          for (var i = 0; i < indicesToProcess.length; i++) {
            var idx = indicesToProcess[i];
            var slide = slides.items[idx];
            var shapes = [];
            for (var j = 0; j < slide.shapes.items.length; j++) {
              var s = slide.shapes.items[j];
              var info = { name: s.name, type: s.type, id: s.id };
              try {
                s.textFrame.load("textRange");
                await context.sync();
                info.text = s.textFrame.textRange.text;
              } catch (e) {}
              shapes.push(info);
            }
            var slideData = { index: idx, id: slide.id, shapeCount: shapes.length, shapes: shapes };
            ${
              withImages
                ? `var img = slide.getImageAsBase64({ width: ${width} });
            await context.sync();
            slideData.imageBase64 = img.value;`
                : ''
            }
            output.push(slideData);
          }
          return { slideCount: slides.items.length, slideWidth: ps.slideWidth, slideHeight: ps.slideHeight, slides: output };
        `
      const target = pool.resolveTarget(presentationId)
      const result = (await pool.sendCommand('executeCode', { code }, target.ws, 120_000)) as {
        slideCount: number
        slideWidth: number
        slideHeight: number
        slides: Array<{
          index: number
          id: string
          shapeCount: number
          shapes: Array<{ name: string; type: string; id: string; text?: string }>
          imageBase64?: string
        }>
      }
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())

      // Build interleaved content blocks
      const content: Array<{ type: 'text'; text: string } | { type: 'image'; data: string; mimeType: string }> = []
      const showing = result.slides.length
      const header = `Deck overview: ${result.slideCount} total slides (${result.slideWidth} x ${result.slideHeight} pt), showing ${showing}${warning ?? ''}`
      content.push({ type: 'text' as const, text: header })

      for (const slide of result.slides) {
        if (slide.imageBase64) {
          content.push({ type: 'image' as const, data: slide.imageBase64, mimeType: 'image/png' })
        }
        const textParts = slide.shapes.filter((s) => s.text).map((s) => s.text!)
        const shapeText = textParts.length > 0 ? `\n${textParts.join('\n')}` : '\n(no text content)'
        content.push({
          type: 'text' as const,
          text: `--- Slide ${slide.index} | ${slide.shapeCount} shapes ---${shapeText}`,
        })
      }

      return { content }
    }),
  )
}
