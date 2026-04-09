import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { z } from 'zod'
import type { ConnectionPool } from './bridge.ts'
import { buildChartRelationship, buildChartXml, buildGraphicFrame, resolveChartPosition } from './chart-builder.ts'
import { recolorSvg, searchIcons } from './icons.ts'
import { buildNotesInjection, readNotesFromDeck } from './notes-helpers.ts'
import type { ThemeInfo } from './xml-helpers.ts'
import {
  autoRegisterContentTypes,
  escapeXml,
  exportSlide,
  extractDeckText,
  extractLayoutsFromZip,
  extractParagraphs,
  extractSlideXmlFromZip,
  extractThemeFromZip,
  extractZipFiles,
  findShapeById,
  listZipPaths,
  NS_A,
  NS_P,
  parseSlideXml,
  reimportSlide,
  replaceParagraphs,
  replaceShape,
  serializeXml,
  updateSlideXmlInZip,
  updateZipFiles,
} from './xml-helpers.ts'

// ---------------------------------------------------------------------------
// Local copy cache: presentationId → { localPath, revision }
// ---------------------------------------------------------------------------

export const localCopyCache = new Map<string, { localPath: string; revision: number }>()
export const themeCache = new Map<string, ThemeInfo>()

// ---------------------------------------------------------------------------
// Concurrent access warning helper
// ---------------------------------------------------------------------------

const sessionConcurrentWarnings = new Map<string, Set<string>>()

export function getConcurrentWarning(
  mcpSessionId: string | undefined,
  presentationId: string,
  activeSessions: number,
): string | null {
  if (!mcpSessionId) return null
  if (activeSessions <= 1) return null

  const warned = sessionConcurrentWarnings.get(mcpSessionId)
  if (warned?.has(presentationId)) return null

  if (!warned) {
    sessionConcurrentWarnings.set(mcpSessionId, new Set([presentationId]))
  } else {
    warned.add(presentationId)
  }

  return '\n\nNote: Other MCP sessions are also connected to the bridge. If they target this presentation, changes apply immediately (last-write-wins).'
}

export function clearSessionWarnings(sessionId: string): void {
  sessionConcurrentWarnings.delete(sessionId)
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Parse a slide range string like "0-3,5,8-10" into a sorted, deduplicated
 * array of zero-based indices: [0,1,2,3,5,8,9,10].
 * Returns null for undefined/empty input (meaning "all slides").
 */
export function parseSlideRange(range: string | undefined): number[] | null {
  if (!range) return null
  const indices = new Set<number>()
  for (const part of range.split(',')) {
    const trimmed = part.trim()
    if (!trimmed) continue
    const dashIdx = trimmed.indexOf('-', 1)
    if (dashIdx === -1) {
      const n = Number(trimmed)
      if (!Number.isInteger(n) || n < 0) throw new Error(`Invalid slide index: "${trimmed}"`)
      indices.add(n)
    } else {
      const start = Number(trimmed.slice(0, dashIdx))
      const end = Number(trimmed.slice(dashIdx + 1))
      if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start) {
        throw new Error(`Invalid slide range: "${trimmed}"`)
      }
      for (let i = start; i <= end; i++) indices.add(i)
    }
  }
  if (indices.size === 0) return null
  return [...indices].sort((a, b) => a - b)
}

// ---------------------------------------------------------------------------
// Tool registration
// ---------------------------------------------------------------------------

export function registerTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  // Helper: get a local file path for a presentation (reused by get_local_copy and inspect_layouts)
  async function getLocalCopyPath(
    connPool: ConnectionPool,
    target: { filePath: string | null; presentationId: string; ws: import('ws').WebSocket },
  ): Promise<string> {
    const filePath = target.filePath

    // Local file — already on disk
    if (filePath && !filePath.startsWith('http')) {
      if (!existsSync(filePath)) throw new Error(`Local file not found: ${filePath}`)
      return filePath
    }

    // Cloud file — check revision for cache validity
    const revCode = `
      var p = context.presentation.properties;
      p.load("revisionNumber");
      await context.sync();
      return p.revisionNumber;
    `
    const currentRevision = (await connPool.sendCommand('executeCode', { code: revCode }, target.ws)) as number

    const cached = localCopyCache.get(target.presentationId)
    if (cached && cached.revision === currentRevision && existsSync(cached.localPath)) {
      return cached.localPath
    }

    // Export fresh copy via Common API getFileAsync
    const exportCode = `
      return new Promise(function(resolve, reject) {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function(result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error.message));
            return;
          }
          var file = result.value;
          var sliceCount = file.sliceCount;
          var sliceData = [];
          var totalSize = 0;
          function getNextSlice(index) {
            if (index >= sliceCount) {
              file.closeAsync();
              var combined = new Uint8Array(totalSize);
              var offset = 0;
              for (var i = 0; i < sliceData.length; i++) {
                var arr = new Uint8Array(sliceData[i]);
                combined.set(arr, offset);
                offset += arr.length;
              }
              var binary = '';
              var chunk = 8192;
              for (var j = 0; j < combined.length; j += chunk) {
                binary += String.fromCharCode.apply(null, combined.subarray(j, Math.min(j + chunk, combined.length)));
              }
              resolve(btoa(binary));
              return;
            }
            file.getSliceAsync(index, function(sliceResult) {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                file.closeAsync();
                reject(new Error(sliceResult.error.message));
                return;
              }
              sliceData.push(sliceResult.value.data);
              totalSize += sliceResult.value.data.length;
              getNextSlice(index + 1);
            });
          }
          getNextSlice(0);
        });
      });
    `
    const base64 = (await connPool.sendCommand('executeCode', { code: exportCode }, target.ws, 120_000)) as string

    const filename = filePath
      ? decodeURIComponent(filePath.split('/').pop() || 'presentation.pptx')
      : 'presentation.pptx'
    const dest = join(tmpdir(), `pptbridge-${Date.now()}-${filename}`)
    writeFileSync(dest, Buffer.from(base64, 'base64'))

    localCopyCache.set(target.presentationId, { localPath: dest, revision: currentRevision })
    return dest
  }

  // --- Tool: list_presentations ---
  server.tool(
    'list_presentations',
    'Lists all PowerPoint presentations currently connected to the bridge server. Shows presentation IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the presentationId to pass to other tools when multiple presentations are open.',
    async () => {
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
    },
  )

  // Helper: enumerate layout usage from existing slides via Office.js
  interface LayoutUsage {
    name: string
    id: string
    usedBySlides: number[]
  }
  async function getLayoutUsage(connPool: ConnectionPool, ws: import('ws').WebSocket): Promise<LayoutUsage[]> {
    const code = `
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      for (var i = 0; i < slides.items.length; i++) {
        slides.items[i].layout.load("name,id");
      }
      await context.sync();
      var seen = {};
      var layouts = [];
      for (var i = 0; i < slides.items.length; i++) {
        var l = slides.items[i].layout;
        if (!seen[l.id]) {
          seen[l.id] = true;
          layouts.push({ name: l.name, id: l.id, usedBySlides: [i] });
        } else {
          for (var j = 0; j < layouts.length; j++) {
            if (layouts[j].id === l.id) { layouts[j].usedBySlides.push(i); break; }
          }
        }
      }
      return layouts;
    `
    return (await connPool.sendCommand('executeCode', { code }, ws)) as LayoutUsage[]
  }

  // --- Tool: inspect_deck ---
  server.tool(
    'inspect_deck',
    'Deck overview: slide dimensions, theme (colors + fonts), and all slides with index, ID, and shape count. Use as the first call to understand deck structure. Theme is cached after the first call. For shape details, follow up with scan_slide or inspect_slide on specific slides. For available layouts, use inspect_layouts.',
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ presentationId }) => {
      try {
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: inspect_layouts ---
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
    async ({ fields, usedOnly, presentationId }) => {
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

      try {
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: add_slide ---
  server.tool(
    'add_slide',
    'Add a new slide from a layout at a specific position, with placeholder text pre-filled and shapes renamed to layout names. Use inspect_layouts to find available layout names and placeholder names. Returns the new slide index and its placeholder shapes with semantic names.',
    {
      layoutName: z
        .string()
        .describe(
          'Layout name from inspect_layouts (e.g. "Title and Content", "Section Header"). Case-insensitive match. Layouts starting with "_" are rejected (technical layouts).',
        ),
      position: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          'Zero-based target slide index. The new slide is inserted at this position (existing slides shift right). Default: append at end.',
        ),
      placeholders: z
        .record(z.string(), z.string())
        .optional()
        .describe(
          'Map of layout placeholder name → text content. E.g. {"section_title": "Backup Slides", "section_subtitle": "Evidence & appendix"}. Names must match layout placeholder names from inspect_layouts.',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ layoutName, position, placeholders, presentationId }) => {
      try {
        // Reject technical layouts
        if (layoutName.startsWith('_')) {
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: Layout "${layoutName}" is a technical gallery separator, not a real layout. Use inspect_layouts to find available layouts.`,
              },
            ],
            isError: true,
          }
        }

        const target = pool.resolveTarget(presentationId)

        // Step 1: Get layout placeholder idx→name map from OOXML
        const localPath = await getLocalCopyPath(pool, target)
        const fileData = readFileSync(localPath)
        const zip = await JSZip.loadAsync(fileData)
        const layouts = await extractLayoutsFromZip(zip)
        const targetLower = layoutName.toLowerCase()
        const layoutInfo = layouts.find((l) => l.name.toLowerCase() === targetLower)
        if (!layoutInfo) {
          const available = layouts
            .filter((l) => !l.name.startsWith('_'))
            .map((l) => l.name)
            .join(', ')
          return {
            content: [
              {
                type: 'text' as const,
                text: `Error: Layout "${layoutName}" not found. Available: ${available}`,
              },
            ],
            isError: true,
          }
        }

        // Build idx→name map from layout placeholders
        const idxToName = new Map<string, string>()
        for (const ph of layoutInfo.placeholders) {
          if (ph.idx !== undefined && ph.name) {
            idxToName.set(String(ph.idx), ph.name)
          }
        }

        // Warn about unknown placeholder names in the input
        const warnings: string[] = []
        if (placeholders) {
          const layoutNames = new Set(layoutInfo.placeholders.map((ph) => ph.name).filter(Boolean))
          for (const key of Object.keys(placeholders)) {
            if (!layoutNames.has(key)) {
              warnings.push(`Placeholder "${key}" not found in layout "${layoutInfo.name}"`)
            }
          }
        }

        // Step 2: Add slide via Office.js (find layout, add, moveTo)
        const addCode = `
          var masters = context.presentation.slideMasters;
          masters.load("items");
          await context.sync();
          var master = masters.items[masters.items.length - 1];
          master.layouts.load("items/id,items/name");
          await context.sync();
          var targetName = ${JSON.stringify(layoutName)}.toLowerCase();
          var layout = null;
          for (var i = 0; i < master.layouts.items.length; i++) {
            if (master.layouts.items[i].name.toLowerCase() === targetName) {
              layout = master.layouts.items[i];
              break;
            }
          }
          if (!layout) {
            var names = master.layouts.items.map(function(l) { return l.name; });
            throw new Error("Layout not found. Available: " + names.join(", "));
          }
          var slides = context.presentation.slides;
          slides.add({ layoutId: layout.id });
          await context.sync();
          slides.load("items");
          await context.sync();
          var newSlide = slides.items[slides.items.length - 1];
          var targetPos = ${position !== undefined ? position : 'undefined'};
          if (targetPos !== undefined) {
            newSlide.moveTo(targetPos);
            await context.sync();
          }
          slides.load("items");
          await context.sync();
          var finalIndex = -1;
          for (var k = 0; k < slides.items.length; k++) {
            if (slides.items[k].id === newSlide.id) { finalIndex = k; break; }
          }
          return { slideIndex: finalIndex, slideId: newSlide.id, slideCount: slides.items.length, layoutName: layout.name };
        `
        const addResult = (await pool.sendCommand('executeCode', { code: addCode }, target.ws)) as {
          slideIndex: number
          slideId: string
          slideCount: number
          layoutName: string
        }

        // Step 3: Export slide XML to get shape id → ph idx mapping
        const exported = await exportSlide(pool, addResult.slideIndex, target.ws)
        const { xmlString } = await extractSlideXmlFromZip(exported.base64)
        const doc = parseSlideXml(xmlString)

        // Parse XML: for each shape with <p:ph idx="N">, record shape id → idx
        const shapeIdToIdx = new Map<string, string>()
        const spElements = doc.getElementsByTagNameNS(NS_P, 'sp')
        for (let i = 0; i < spElements.length; i++) {
          const sp = spElements[i]!
          const nvSpPr = sp.getElementsByTagNameNS(NS_P, 'nvSpPr')[0]
          if (!nvSpPr) continue
          const cNvPr = nvSpPr.getElementsByTagNameNS(NS_P, 'cNvPr')[0]
          const nvPr = nvSpPr.getElementsByTagNameNS(NS_P, 'nvPr')[0]
          if (!cNvPr || !nvPr) continue
          const phEl = nvPr.getElementsByTagNameNS(NS_P, 'ph')[0]
          if (!phEl) continue
          const idx = phEl.getAttribute('idx')
          const id = cNvPr.getAttribute('id')
          if (idx && id) {
            shapeIdToIdx.set(id, idx)
          }
        }

        // Build rename map: Office.js shape id → layout semantic name
        // Also build text map for placeholders to fill
        const renameMap: Record<string, string> = {}
        const textMap: Record<string, string> = {}
        for (const [shapeId, idx] of shapeIdToIdx) {
          const semanticName = idxToName.get(idx)
          if (semanticName) {
            renameMap[shapeId] = semanticName
            if (placeholders?.[semanticName]) {
              textMap[shapeId] = placeholders[semanticName]
            }
          }
        }

        // Step 4: Rename shapes and set text via Office.js
        const renameCode = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          var slide = slides.items[${addResult.slideIndex}];
          slide.shapes.load("items/id,items/name");
          await context.sync();
          var renameMap = ${JSON.stringify(renameMap)};
          var textMap = ${JSON.stringify(textMap)};
          var results = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var newName = renameMap[s.id];
            if (newName) {
              s.name = newName;
            }
            var text = textMap[s.id];
            if (text !== undefined) {
              try {
                s.textFrame.textRange.text = text;
              } catch(e) {}
            }
          }
          await context.sync();
          // Collect final info for renamed shapes (layout placeholders)
          slide.shapes.load("items/id,items/name");
          await context.sync();
          var placeholders = [];
          for (var j = 0; j < slide.shapes.items.length; j++) {
            var s = slide.shapes.items[j];
            if (!renameMap[s.id]) continue;
            var info = { id: s.id, name: s.name, text: "" };
            try {
              var tf = s.getTextFrameOrNullObject();
              tf.load(["hasText", "textRange"]);
              await context.sync();
              if (!tf.isNullObject && tf.hasText) info.text = tf.textRange.text;
            } catch(e) {}
            placeholders.push(info);
          }
          return placeholders;
        `
        const phResult = (await pool.sendCommand('executeCode', { code: renameCode }, target.ws)) as Array<{
          id: string
          name: string
          text: string
        }>

        const result: Record<string, unknown> = {
          slideIndex: addResult.slideIndex,
          slideCount: addResult.slideCount,
          layoutName: addResult.layoutName,
          placeholders: phResult,
        }
        if (warnings.length > 0) {
          result.warnings = warnings
        }

        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: inspect_slide ---
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
    async ({ slideRange, presentationId }) => {
      try {
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: scan_slide ---
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
    async ({ slideRange, namePattern, shapeType, presentationId }) => {
      try {
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
          const nameRegex = namePattern ? new RegExp(`^${namePattern.replace(/\*/g, '.*')}$`, 'i') : null
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: screenshot_slide ---
  server.tool(
    'screenshot_slide',
    'Slide screenshot (~1000 tokens): captures one slide as PNG image. Use to visually verify layout after changes. Do NOT loop over all slides — use preview_deck instead.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from list_slides results'),
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
    async ({ slideIndex, width, height, presentationId }) => {
      try {
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
        const hint =
          message.includes('getImageAsBase64') || message.includes('not a function')
            ? ' (This API requires PowerPoint 16.96+ with PowerPointApi 1.8 support)'
            : ''
        return { content: [{ type: 'text' as const, text: `Error: ${message}${hint}` }], isError: true }
      }
    },
  )

  // --- Tool: copy_slides ---
  server.tool(
    'copy_slides',
    'Copies slides between two open presentations entirely server-side — the Base64 data never enters Claude context. Exports from the source presentation and inserts into the destination in a single operation. Both presentations must be connected to the bridge. Requires PowerPointApi 1.8.',
    {
      sourceSlideIndex: z.number().int().min(0).describe('Zero-based slide index to copy from the source presentation'),
      sourcePresentationId: z.string().describe('Source presentation ID from list_presentations'),
      destinationPresentationId: z.string().describe('Destination presentation ID from list_presentations'),
      targetSlideId: z
        .string()
        .optional()
        .describe(
          'Insert after this slide ID in destination (format: "nnn#" or "#mmmmmmmmm" or "nnn#mmmmmmmmm"). If omitted, inserts at the beginning.',
        ),
      formatting: z
        .enum(['KeepSourceFormatting', 'UseDestinationTheme'])
        .optional()
        .describe('Formatting mode. Default: KeepSourceFormatting.'),
    },
    async ({ sourceSlideIndex, sourcePresentationId, destinationPresentationId, targetSlideId, formatting }) => {
      try {
        // Step 1: Export slide from source presentation
        const source = pool.resolveTarget(sourcePresentationId)
        const exportCode = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${sourceSlideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${sourceSlideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${sourceSlideIndex}];
          var result = slide.exportAsBase64();
          await context.sync();
          return { base64: result.value, slideIndex: ${sourceSlideIndex}, slideId: slide.id };
        `
        const exported = (await pool.sendCommand('executeCode', { code: exportCode }, source.ws)) as {
          base64: string
          slideIndex: number
          slideId: string
        }

        // Step 2: Insert into destination presentation
        const dest = pool.resolveTarget(destinationPresentationId)
        const optionsParts: string[] = []
        if (formatting) {
          optionsParts.push(`formatting: "${formatting}"`)
        }
        if (targetSlideId) {
          optionsParts.push(`targetSlideId: "${targetSlideId}"`)
        }
        const optionsArg = optionsParts.length > 0 ? `, { ${optionsParts.join(', ')} }` : ''

        const insertCode = `
          context.presentation.insertSlidesFromBase64("${exported.base64}"${optionsArg});
          await context.sync();
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return { slideCount: slides.items.length };
        `
        const inserted = (await pool.sendCommand('executeCode', { code: insertCode }, dest.ws)) as {
          slideCount: number
        }

        const warning = getConcurrentWarning(getSessionId(), dest.presentationId, getActiveSessionCount())
        const text =
          JSON.stringify(
            {
              copied: { slideIndex: exported.slideIndex, slideId: exported.slideId },
              destination: { slideCount: inserted.slideCount },
            },
            null,
            2,
          ) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: insert_image ---
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
    async ({ source, sourceType, slideIndex, left, top, width, height, color, presentationId }) => {
      try {
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: preview_deck ---
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
    async ({ slideRange, imageWidth, includeImages, presentationId }) => {
      try {
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
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: get_local_copy ---
  server.tool(
    'get_local_copy',
    'Returns a local file path for the presentation. For local files, returns the existing path. For SharePoint/cloud files, exports server-side and saves to a temp .pptx. Caches by revision number — re-exports only when the presentation has been saved since last export.',
    {
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const cachedBefore = localCopyCache.get(target.presentationId)?.localPath
        const localPath = await getLocalCopyPath(pool, target)
        const isLocal = target.filePath && !target.filePath.startsWith('http')
        const cached = localCopyCache.get(target.presentationId)
        const source = isLocal ? 'local' : cachedBefore === localPath ? 'cached' : 'exported'
        const result: Record<string, unknown> = { localPath, source }
        if (cached) result.revision = cached.revision
        return { content: [{ type: 'text' as const, text: JSON.stringify(result) }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: read_shape_paragraphs ---
  server.tool(
    'read_shape_paragraphs',
    "Read raw OOXML <a:p> paragraphs from a shape's text body. Returns the paragraph XML as a string — preserves all formatting (bold, colors, bullets, etc.) that textRange.text strips. Use with the /pptx skill's OOXML knowledge to understand and modify the XML.",
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from list_slides results'),
      shapeId: z.string().describe('Shape ID from inspect_slide results (e.g. "5")'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, shapeId, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { xmlString } = await extractSlideXmlFromZip(exported.base64)
        const doc = parseSlideXml(xmlString)
        const shape = findShapeById(doc, shapeId)
        if (!shape) {
          throw new Error(`Shape with ID "${shapeId}" not found on slide ${slideIndex}`)
        }
        const paragraphXml = extractParagraphs(shape)
        return { content: [{ type: 'text' as const, text: paragraphXml }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: read_deck_text ---
  server.tool(
    'read_deck_text',
    'Lightweight text extractor: returns slide titles and body text as plain strings (~20x smaller than inspect_slide). Use for content review, narrative analysis, or any read-only text task. Supports slideRange and optional speaker notes.',
    {
      slideRange: z.string().optional().describe('Slide range, e.g. "0-5", "2,4,7". Zero-based. Omit for all slides.'),
      includeNotes: z.boolean().optional().describe('Include speaker notes. Default: false.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideRange, includeNotes, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const localPath = await getLocalCopyPath(pool, target)
        const zipBuffer = readFileSync(localPath)
        const indices = slideRange ? parseSlideRange(slideRange) : null
        const result = await extractDeckText(zipBuffer, indices, includeNotes === true)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: edit_shape_paragraphs ---
  server.tool(
    'edit_shape_paragraphs',
    "Replace paragraph content of a shape with raw OOXML <a:p> XML. Preserves <a:bodyPr> and <a:lstStyle>. Use read_shape_paragraphs first to get the current XML, modify it (using /pptx skill knowledge), then write it back. The slide is exported, modified server-side, and reimported — data never enters Claude's context.",
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      shapeId: z.string().describe('Shape ID from inspect_slide results'),
      xml: z.string().describe('The <a:p> paragraph XML to replace the current text body content with'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, shapeId, xml, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { zip, xmlString } = await extractSlideXmlFromZip(exported.base64)
        const doc = parseSlideXml(xmlString)
        const shape = findShapeById(doc, shapeId)
        if (!shape) {
          throw new Error(`Shape with ID "${shapeId}" not found on slide ${slideIndex}`)
        }
        replaceParagraphs(doc, shape, xml)
        const modifiedBase64 = await updateSlideXmlInZip(zip, serializeXml(doc))
        await reimportSlide(pool, modifiedBase64, exported.slideId, exported.prevSlideId, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify({ success: true }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: read_slide_xml ---
  server.tool(
    'read_slide_xml',
    "Read the full raw OOXML of a slide, or filter to a specific shape. Returns the slide's ppt/slides/slide1.xml content. Use with the /pptx skill's OOXML knowledge to understand the XML structure.",
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from list_slides results'),
      shapeId: z
        .string()
        .optional()
        .describe("Optional shape ID to filter to. If provided, returns only that shape's <p:sp> element."),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, shapeId, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { xmlString } = await extractSlideXmlFromZip(exported.base64)

        if (shapeId) {
          const doc = parseSlideXml(xmlString)
          const shape = findShapeById(doc, shapeId)
          if (!shape) {
            throw new Error(`Shape with ID "${shapeId}" not found on slide ${slideIndex}`)
          }
          return { content: [{ type: 'text' as const, text: serializeXml(shape) }] }
        }

        return { content: [{ type: 'text' as const, text: xmlString }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: edit_slide_xml ---
  server.tool(
    'edit_slide_xml',
    "Edit slide XML and reimport. Two modes: (1) xml mode — provide finished XML string (use read_slide_xml first to get current XML), (2) code mode — provide JS code that manipulates the pre-parsed DOM (receives: doc, findShapeById, NS_P, NS_A, escapeXml, serializeXml, DOMParser). Code mode preserves untouched attributes. The slide is exported, modified server-side, and reimported — data never enters Claude's context.",
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      xml: z
        .string()
        .optional()
        .describe(
          "Modified XML — full slide XML or a single shape's <p:sp> element (when shapeId is provided). Mutually exclusive with 'code'.",
        ),
      code: z
        .string()
        .optional()
        .describe(
          "JS code that manipulates the pre-parsed slide DOM. Receives: doc (Document), findShapeById(id) → Element|null, NS_P, NS_A (namespace strings), escapeXml(text), serializeXml(node), DOMParser. Mutually exclusive with 'xml'.",
        ),
      explanation: z
        .string()
        .optional()
        .describe('Brief description of what the code does (for logging, max 50 chars). Only used with code mode.'),
      shapeId: z
        .string()
        .optional()
        .describe(
          "Optional shape ID (xml mode only). If provided, replaces only that shape's <p:sp> element instead of the full slide XML.",
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, xml, code, shapeId, presentationId }) => {
      if ((!xml && !code) || (xml && code)) {
        return {
          content: [
            { type: 'text' as const, text: "Error: Provide either 'xml' or 'code', not both and not neither." },
          ],
          isError: true,
        }
      }

      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { zip, xmlString } = await extractSlideXmlFromZip(exported.base64)

        let finalXml: string
        if (code) {
          // Code mode: run agent-provided JS against the pre-parsed DOM
          const doc = parseSlideXml(xmlString)
          const sandbox = {
            doc,
            findShapeById: (id: string) => findShapeById(doc, id),
            NS_P,
            NS_A,
            escapeXml,
            serializeXml: (node: Document | Element) => serializeXml(node),
            DOMParser,
          }
          try {
            const keys = Object.keys(sandbox)
            const values = Object.values(sandbox)
            const fn = new Function(...keys, code)
            fn(...values)
          } catch (codeErr: unknown) {
            const msg = codeErr instanceof Error ? codeErr.message : String(codeErr)
            throw new Error(`Code execution error: ${msg}`)
          }
          finalXml = serializeXml(doc)
        } else if (shapeId) {
          const doc = parseSlideXml(xmlString)
          const shape = findShapeById(doc, shapeId)
          if (!shape) {
            throw new Error(`Shape with ID "${shapeId}" not found on slide ${slideIndex}`)
          }
          replaceShape(doc, shape, xml!)
          finalXml = serializeXml(doc)
        } else {
          finalXml = xml!
        }

        const modifiedBase64 = await updateSlideXmlInZip(zip, finalXml)
        await reimportSlide(pool, modifiedBase64, exported.slideId, exported.prevSlideId, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify({ success: true }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: duplicate_slide ---
  server.tool(
    'duplicate_slide',
    'Duplicate a slide within the same presentation. Exports the slide and reimports it at the specified position. Data stays server-side.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based index of the slide to duplicate'),
      insertAfter: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          'Zero-based slide index to insert the duplicate after. Default: same as slideIndex (duplicate appears right after the source).',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, insertAfter, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const insertPos = insertAfter ?? slideIndex

        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          if (${insertPos} >= slides.items.length) {
            throw new Error("insertAfter index " + ${insertPos} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          var result = slide.exportAsBase64();
          await context.sync();
          var targetId = slides.items[${insertPos}].id;
          context.presentation.insertSlidesFromBase64(result.value, {
            formatting: "KeepSourceFormatting",
            targetSlideId: targetId
          });
          await context.sync();
          slides.load("items");
          await context.sync();
          return { duplicatedSlideIndex: ${slideIndex}, insertedAfter: ${insertPos}, slideCount: slides.items.length };
        `
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: verify_slides ---
  server.tool(
    'verify_slides',
    'Run programmatic checks on a slide: detect overlapping shapes, out-of-bounds shapes, empty text, tiny shapes, unused placeholders, placeholder drift from layout defaults, and full-bleed background covers. Returns a list of issues found. Uses the same shape data as inspect_slide — no OOXML needed.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      checks: z
        .array(
          z.enum([
            'overlap',
            'bounds',
            'empty_text',
            'tiny_shapes',
            'unused_placeholder',
            'layout_drift',
            'background_cover',
          ]),
        )
        .optional()
        .describe('Checks to run. Default: all checks.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, checks, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const enabledChecks = checks ?? [
          'overlap',
          'bounds',
          'empty_text',
          'tiny_shapes',
          'unused_placeholder',
          'layout_drift',
          'background_cover',
        ]

        const checkLayoutDrift = enabledChecks.includes('layout_drift')

        // Reuse inspect_slide's shape-loading logic
        const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          slide.shapes.load("items");
          await context.sync();
          var shapes = [];
          var placeholderTypes = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var info = {
              name: s.name,
              id: s.id,
              left: s.left,
              top: s.top,
              width: s.width,
              height: s.height
            };
            try {
              var tf = s.getTextFrameOrNullObject();
              tf.load(["hasText", "textRange"]);
              await context.sync();
              if (!tf.isNullObject) {
                info.text = tf.hasText ? tf.textRange.text : "";
                info.hasText = tf.hasText;
              }
            } catch (e) {}
            if (s.type === "Placeholder") {
              try {
                var pf = s.placeholderFormat;
                pf.load("type");
                await context.sync();
                info.isPlaceholder = true;
                info.placeholderType = pf.type;
                placeholderTypes.push({ shapeIndex: i, type: pf.type });
              } catch (e) {}
            }
            shapes.push(info);
          }

          // Conditionally load layout placeholder positions for drift check
          var layoutMap = {};
          if (${checkLayoutDrift} && placeholderTypes.length > 0) {
            try {
              var layout = slide.layout;
              layout.load("name");
              var layoutShapes = layout.shapes;
              layoutShapes.load("items");
              await context.sync();
              for (var li = 0; li < layoutShapes.items.length; li++) {
                var ls = layoutShapes.items[li];
                if (ls.type !== "Placeholder") continue;
                try {
                  var lph = ls.placeholderFormat;
                  lph.load("type");
                  ls.load("left,top,width,height,name");
                  await context.sync();
                  layoutMap[lph.type] = {
                    name: ls.name,
                    left: ls.left,
                    top: ls.top,
                    width: ls.width,
                    height: ls.height
                  };
                } catch (e) {}
              }
            } catch (e) {}
            // Attach layout match to shapes
            for (var pi = 0; pi < placeholderTypes.length; pi++) {
              var pt = placeholderTypes[pi];
              var match = layoutMap[pt.type];
              if (match) {
                shapes[pt.shapeIndex].layoutMatch = match;
              }
            }
          }

          // Also get slide dimensions
          var ps = context.presentation.pageSetup;
          ps.load("slideWidth,slideHeight");
          await context.sync();
          return { shapes: shapes, slideWidth: ps.slideWidth, slideHeight: ps.slideHeight };
        `
        const slideData = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
          shapes: Array<{
            name: string
            id: string
            left: number
            top: number
            width: number
            height: number
            text?: string
            hasText?: boolean
            isPlaceholder?: boolean
            placeholderType?: string
            layoutMatch?: { name: string; left: number; top: number; width: number; height: number }
          }>
          slideWidth: number
          slideHeight: number
        }

        const issues: Array<{
          type: string
          severity: 'warning' | 'error'
          shapeIds: string[]
          description: string
        }> = []

        const { shapes, slideWidth, slideHeight } = slideData

        // Overlap check: AABB collision
        if (enabledChecks.includes('overlap')) {
          for (let i = 0; i < shapes.length; i++) {
            for (let j = i + 1; j < shapes.length; j++) {
              const a = shapes[i]!
              const b = shapes[j]!
              if (
                a.left < b.left + b.width &&
                a.left + a.width > b.left &&
                a.top < b.top + b.height &&
                a.top + a.height > b.top
              ) {
                issues.push({
                  type: 'overlap',
                  severity: 'warning',
                  shapeIds: [a.id, b.id],
                  description: `"${a.name}" and "${b.name}" overlap`,
                })
              }
            }
          }
        }

        // Bounds check: shape extends beyond slide
        if (enabledChecks.includes('bounds')) {
          for (const s of shapes) {
            const outOfBounds: string[] = []
            if (s.left < 0) outOfBounds.push('left of slide')
            if (s.top < 0) outOfBounds.push('above slide')
            if (s.left + s.width > slideWidth) outOfBounds.push('right of slide')
            if (s.top + s.height > slideHeight) outOfBounds.push('below slide')
            if (outOfBounds.length > 0) {
              issues.push({
                type: 'bounds',
                severity: 'warning',
                shapeIds: [s.id],
                description: `"${s.name}" extends ${outOfBounds.join(', ')}`,
              })
            }
          }
        }

        // Empty text check
        if (enabledChecks.includes('empty_text')) {
          for (const s of shapes) {
            if (s.text !== undefined && s.text.trim() === '') {
              issues.push({
                type: 'empty_text',
                severity: 'warning',
                shapeIds: [s.id],
                description: `"${s.name}" has an empty text frame`,
              })
            }
          }
        }

        // Tiny shapes check
        if (enabledChecks.includes('tiny_shapes')) {
          for (const s of shapes) {
            if (s.width < 10 || s.height < 10) {
              issues.push({
                type: 'tiny_shapes',
                severity: 'warning',
                shapeIds: [s.id],
                description: `"${s.name}" is very small (${s.width.toFixed(1)} x ${s.height.toFixed(1)} pt)`,
              })
            }
          }
        }

        // Unused placeholder check: placeholder shapes with no text content
        if (enabledChecks.includes('unused_placeholder')) {
          for (const s of shapes) {
            if (s.isPlaceholder && !s.hasText) {
              issues.push({
                type: 'unused_placeholder',
                severity: 'warning',
                shapeIds: [s.id],
                description: `"${s.name}" is an unused placeholder — delete it or fill it with content`,
              })
            }
          }
        }

        // Layout drift check: placeholder position vs layout default
        if (checkLayoutDrift) {
          const DRIFT_THRESHOLD = 2 // points
          for (const s of shapes) {
            if (!s.isPlaceholder || !s.layoutMatch) continue
            const lm = s.layoutMatch
            const drifts: string[] = []
            if (Math.abs(s.left - lm.left) > DRIFT_THRESHOLD) drifts.push(`left: ${s.left} vs layout ${lm.left}`)
            if (Math.abs(s.top - lm.top) > DRIFT_THRESHOLD) drifts.push(`top: ${s.top} vs layout ${lm.top}`)
            if (Math.abs(s.width - lm.width) > DRIFT_THRESHOLD) drifts.push(`width: ${s.width} vs layout ${lm.width}`)
            if (Math.abs(s.height - lm.height) > DRIFT_THRESHOLD)
              drifts.push(`height: ${s.height} vs layout ${lm.height}`)
            if (drifts.length > 0) {
              issues.push({
                type: 'layout_drift',
                severity: 'warning',
                shapeIds: [s.id],
                description: `"${s.name}" drifted from layout: ${drifts.join(', ')}`,
              })
            }
          }
        }

        // Background cover check: non-placeholder shapes covering most of the slide
        if (enabledChecks.includes('background_cover')) {
          const dimThreshold = 0.85
          const areaThreshold = 0.9
          const slideArea = slideWidth * slideHeight
          for (const s of shapes) {
            if (s.isPlaceholder) continue
            const widthRatio = s.width / slideWidth
            const heightRatio = s.height / slideHeight
            if (
              widthRatio >= dimThreshold &&
              heightRatio >= dimThreshold &&
              s.width * s.height >= slideArea * areaThreshold
            ) {
              issues.push({
                type: 'background_cover',
                severity: 'error',
                shapeIds: [s.id],
                description: `"${s.name}" covers ${(widthRatio * 100).toFixed(0)}% x ${(heightRatio * 100).toFixed(0)}% of the slide — this destroys the layout background, logo, and design system. Delete this shape and use the layout's background instead.`,
              })
            }
          }
        }

        const result = { slideIndex, shapeCount: shapes.length, issueCount: issues.length, issues }
        return { content: [{ type: 'text' as const, text: JSON.stringify(result, null, 2) }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: read_slide_zip ---
  server.tool(
    'read_slide_zip',
    'Read multiple files from the exported slide zip. Returns slide XML, relationships, chart XMLs, and Content_Types. Use this to inspect chart data, rels, or other zip contents beyond what read_slide_xml provides. When no paths specified, auto-discovers all text/XML files in the zip.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      paths: z
        .array(z.string())
        .optional()
        .describe(
          'Specific zip paths to read (e.g. ["ppt/slides/slide1.xml", "ppt/charts/chart1.xml"]). If omitted, auto-discovers all text/XML files.',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, paths, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { zip, files } = await extractZipFiles(exported.base64, paths)
        const allPaths = listZipPaths(zip)
        const result = { zipContents: files, allPaths }
        return { content: [{ type: 'text' as const, text: JSON.stringify(result, null, 2) }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: edit_slide_zip ---
  server.tool(
    'edit_slide_zip',
    'Update multiple files in the slide zip and reimport in a single operation. Accepts a map of { path: content } — can modify existing files or add new ones (e.g. chart XML + rels). Auto-registers Content_Types for new chart files. Use read_slide_zip first to get the current content.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      files: z
        .record(z.string(), z.string())
        .describe(
          'Map of { zipPath: newContent }. Can include existing paths (to modify) or new paths (to add). Example: { "ppt/slides/slide1.xml": "<p:sld>...</p:sld>", "ppt/charts/chart1.xml": "<c:chartSpace>...</c:chartSpace>" }',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, files, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { zip } = await extractZipFiles(exported.base64)

        // Detect new files (not already in zip) for Content_Types auto-registration
        const existingPaths = new Set(listZipPaths(zip))
        const newPaths = Object.keys(files).filter((p) => !existingPaths.has(p))

        // Apply user-provided file changes first (explicit takes precedence)
        const modifiedBase64 = await updateZipFiles(zip, files)

        // Auto-register Content_Types for new chart files (if not already handled by user)
        if (newPaths.length > 0 && !files['[Content_Types].xml']) {
          const { zip: updatedZip } = await extractZipFiles(modifiedBase64)
          await autoRegisterContentTypes(updatedZip, newPaths)
          const finalBase64 = await updatedZip.generateAsync({ type: 'base64' })
          await reimportSlide(pool, finalBase64, exported.slideId, exported.prevSlideId, target.ws)
        } else {
          await reimportSlide(pool, modifiedBase64, exported.slideId, exported.prevSlideId, target.ws)
        }

        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text =
          JSON.stringify({ success: true, filesUpdated: Object.keys(files).length, newFiles: newPaths }, null, 2) +
          (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: edit_slide_chart ---
  server.tool(
    'edit_slide_chart',
    'Create a chart on a slide from structured data. Generates all OOXML automatically (chart XML, rels, graphic frame, Content_Types). Supports column, bar, line, pie, doughnut, and area charts with multiple series.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      chartType: z.enum(['column', 'bar', 'line', 'pie', 'doughnut', 'area']).describe('Chart type'),
      title: z.string().describe('Chart title'),
      categories: z.array(z.string()).describe('Category labels (x-axis or pie slices)'),
      series: z
        .array(
          z.object({
            name: z.string().describe('Series name'),
            values: z.array(z.number()).describe('Data values (one per category)'),
          }),
        )
        .min(1)
        .describe('Data series'),
      position: z
        .object({
          left: z.number().optional().describe('Left position in points'),
          top: z.number().optional().describe('Top position in points'),
          width: z.number().optional().describe('Width in points'),
          height: z.number().optional().describe('Height in points'),
        })
        .optional()
        .describe('Chart position in points. Defaults to centered on slide.'),
      options: z
        .object({
          stacked: z.boolean().optional().describe('Use stacked grouping (bar/column/line/area)'),
          showDataLabels: z.boolean().optional().describe('Show data labels (default true)'),
          showLegend: z.boolean().optional().describe('Show legend (default true)'),
          legendPosition: z
            .enum(['t', 'b', 'l', 'r'])
            .optional()
            .describe('Legend position: t=top, b=bottom, l=left, r=right'),
        })
        .optional()
        .describe('Chart options'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, chartType, title, categories, series, position, options, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)

        // 1. Export the slide
        const exported = await exportSlide(pool, slideIndex, target.ws)
        const { zip } = await extractZipFiles(exported.base64)

        // 2. Determine next chart number (scan existing ppt/charts/ in zip)
        const existingPaths = listZipPaths(zip)
        const chartPaths = existingPaths.filter((p) => p.startsWith('ppt/charts/chart') && p.endsWith('.xml'))
        const chartNums = chartPaths.map((p) => {
          const m = p.match(/chart(\d+)\.xml$/)
          return m ? Number(m[1]) : 0
        })
        const nextChartNum = chartNums.length > 0 ? Math.max(...chartNums) + 1 : 1
        const chartFileName = `chart${nextChartNum}.xml`
        const chartZipPath = `ppt/charts/${chartFileName}`

        // 3. Determine next rId from slide rels
        const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
        const relsContent = zip.file(relsPath)
          ? await zip.file(relsPath)!.async('string')
          : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        const rIdMatches = [...relsContent.matchAll(/Id="rId(\d+)"/g)]
        const rIdNums = rIdMatches.map((m) => Number(m[1]))
        const nextRIdNum = rIdNums.length > 0 ? Math.max(...rIdNums) + 1 : 1
        const rId = `rId${nextRIdNum}`

        // 4. Generate chart XML
        const chartXml = buildChartXml(chartType, title, categories, series, options)

        // 5. Generate graphic frame and inject into slide XML
        const slideXmlPath = 'ppt/slides/slide1.xml'
        const slideXml = await zip.file(slideXmlPath)!.async('string')
        const emuPos = resolveChartPosition(position)

        // Find the highest shape ID in the slide to avoid conflicts
        const shapeIdMatches = [...slideXml.matchAll(/id="(\d+)"/g)]
        const shapeIds = shapeIdMatches.map((m) => Number(m[1]))
        const nextShapeId = shapeIds.length > 0 ? Math.max(...shapeIds) + 1 : 100
        const chartName = `Chart ${nextChartNum}`

        const graphicFrame = buildGraphicFrame(rId, emuPos, chartName, nextShapeId)

        // Inject graphic frame before </p:spTree>
        const modifiedSlideXml = slideXml.replace('</p:spTree>', `${graphicFrame}</p:spTree>`)

        // 6. Add chart relationship to rels
        const relEntry = buildChartRelationship(rId, `../charts/${chartFileName}`)
        const modifiedRels = relsContent.replace('</Relationships>', `${relEntry}</Relationships>`)

        // 7. Write all files and reimport
        const files: Record<string, string> = {
          [slideXmlPath]: modifiedSlideXml,
          [chartZipPath]: chartXml,
          [relsPath]: modifiedRels,
        }

        const newPaths = Object.keys(files).filter((p) => !new Set(existingPaths).has(p))
        const modifiedBase64 = await updateZipFiles(zip, files)

        // Auto-register Content_Types for the new chart file
        if (newPaths.length > 0) {
          const { zip: updatedZip } = await extractZipFiles(modifiedBase64)
          await autoRegisterContentTypes(updatedZip, newPaths)
          const finalBase64 = await updatedZip.generateAsync({ type: 'base64' })
          await reimportSlide(pool, finalBase64, exported.slideId, exported.prevSlideId, target.ws)
        } else {
          await reimportSlide(pool, modifiedBase64, exported.slideId, exported.prevSlideId, target.ws)
        }

        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text =
          JSON.stringify(
            {
              success: true,
              chartType,
              title,
              seriesCount: series.length,
              categoryCount: categories.length,
              chartFile: chartZipPath,
            },
            null,
            2,
          ) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: search_text ---
  server.tool(
    'search_text',
    'Search for text across all slides (or a slide range) in the presentation — like grep for slides. Searches shape text, table cells, and speaker notes. Returns matching shapes with context at the chosen level. Case-insensitive substring by default; supports regex.',
    {
      query: z.string().describe('Text to search for. Plain substring by default, or a regex pattern when regex=true.'),
      slideRange: z
        .string()
        .optional()
        .describe('Optional slide range to search, e.g. "0-4" or "2-7". Zero-based. Omit to search all slides.'),
      caseSensitive: z.boolean().optional().describe('Case-sensitive search. Default: false.'),
      regex: z.boolean().optional().describe('Treat query as a regular expression. Default: false (plain substring).'),
      context: z
        .enum(['shape', 'slide', 'none'])
        .optional()
        .describe(
          'Result detail level. "shape" (default): matching shapes with text. "slide": all shapes on matching slides with matched markers. "none": just matching slide indices.',
        ),
      includeNotes: z
        .boolean()
        .optional()
        .describe('Search speaker notes in addition to slide content. Default: true.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ query, slideRange, caseSensitive, regex, context: contextLevel, includeNotes, presentationId }) => {
      try {
        const cs = caseSensitive === true
        const useRegex = regex === true
        const ctxLevel = contextLevel ?? 'shape'
        const searchNotes = includeNotes !== false
        const escapedQuery = JSON.stringify(query)
        const code = `
          var caseSensitive = ${cs};
          var useRegex = ${useRegex};
          var query = ${escapedQuery};
          var ctxLevel = ${JSON.stringify(ctxLevel)};
          var searchNotes = ${searchNotes};
          var slideRangeStr = ${slideRange ? JSON.stringify(slideRange) : 'null'};

          function testMatch(text) {
            if (useRegex) {
              var flags = caseSensitive ? "" : "i";
              var re = new RegExp(query, flags);
              return re.test(text);
            }
            var a = caseSensitive ? text : text.toLowerCase();
            var b = caseSensitive ? query : query.toLowerCase();
            return a.indexOf(b) !== -1;
          }

          function extractShapeTexts(shapes) {
            var results = [];
            for (var j = 0; j < shapes.items.length; j++) {
              var shape = shapes.items[j];
              var texts = [];
              try {
                shape.textFrame.load("textRange");
                try { shape.textFrame.textRange.load("text"); } catch(e) {}
              } catch (e) {}
              try {
                if (shape.type === "Table") {
                  shape.table.load("rowCount,columnCount");
                }
              } catch (e) {}
            }
            return shapes;
          }

          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          var total = slides.items.length;
          var startIdx = 0;
          var endIdx = total - 1;
          if (slideRangeStr) {
            var parts = slideRangeStr.split("-");
            startIdx = parseInt(parts[0], 10);
            if (parts.length > 1) endIdx = parseInt(parts[1], 10);
            else endIdx = startIdx;
            if (startIdx < 0) startIdx = 0;
            if (endIdx >= total) endIdx = total - 1;
          }

          var slideResults = [];
          for (var si = startIdx; si <= endIdx; si++) {
            var slide = slides.items[si];
            slide.shapes.load("items");
            await context.sync();

            var shapeEntries = [];
            var slideHasMatch = false;

            for (var j = 0; j < slide.shapes.items.length; j++) {
              var shape = slide.shapes.items[j];
              var shapeId = String(shape.id);
              var shapeName = shape.name;
              var shapeType = shape.type;
              var texts = [];

              try {
                shape.textFrame.load("textRange");
                await context.sync();
                var t = shape.textFrame.textRange.text;
                if (t) texts.push({ source: "shape", text: t });
              } catch (e) {}

              if (shapeType === "Table") {
                try {
                  shape.table.load("rowCount,columnCount");
                  await context.sync();
                  var rc = shape.table.rowCount;
                  var cc = shape.table.columnCount;
                  for (var r = 0; r < rc; r++) {
                    for (var c = 0; c < cc; c++) {
                      try {
                        var cell = shape.table.getCell(r, c);
                        cell.body.load("text");
                        await context.sync();
                        if (cell.body.text) texts.push({ source: "tableCell", text: cell.body.text, row: r, col: c });
                      } catch (e2) {}
                    }
                  }
                } catch (e) {}
              }

              var matched = false;
              for (var ti = 0; ti < texts.length; ti++) {
                if (testMatch(texts[ti].text)) { matched = true; break; }
              }

              if (matched) slideHasMatch = true;

              shapeEntries.push({
                shapeId: shapeId,
                shapeName: shapeName,
                shapeType: shapeType,
                matched: matched,
                texts: texts
              });
            }

            var noteText = null;
            var noteMatched = false;
            if (searchNotes) {
              try {
                var ns = slide.notesSlide;
                ns.shapes.load("items");
                await context.sync();
                for (var ni = 0; ni < ns.shapes.items.length; ni++) {
                  try {
                    ns.shapes.items[ni].textFrame.load("textRange");
                    await context.sync();
                    var nt = ns.shapes.items[ni].textFrame.textRange.text;
                    if (nt && nt.trim()) {
                      noteText = (noteText || "") + nt;
                    }
                  } catch (e) {}
                }
                if (noteText && testMatch(noteText)) {
                  noteMatched = true;
                  slideHasMatch = true;
                }
              } catch (e) {}
            }

            if (slideHasMatch) {
              if (ctxLevel === "none") {
                slideResults.push(si);
              } else if (ctxLevel === "slide") {
                var entry = { slideIndex: si, shapes: [] };
                for (var k = 0; k < shapeEntries.length; k++) {
                  var se = shapeEntries[k];
                  var shapeInfo = {
                    shapeId: se.shapeId,
                    shapeName: se.shapeName,
                    matched: se.matched
                  };
                  for (var ti2 = 0; ti2 < se.texts.length; ti2++) {
                    var tx = se.texts[ti2];
                    if (tx.source === "shape") shapeInfo.text = tx.text;
                    else if (tx.source === "tableCell") {
                      if (!shapeInfo.tableCells) shapeInfo.tableCells = [];
                      shapeInfo.tableCells.push({ row: tx.row, col: tx.col, text: tx.text, matched: testMatch(tx.text) });
                    }
                  }
                  entry.shapes.push(shapeInfo);
                }
                if (noteText) entry.notes = { text: noteText, matched: noteMatched };
                slideResults.push(entry);
              } else {
                for (var k2 = 0; k2 < shapeEntries.length; k2++) {
                  var se2 = shapeEntries[k2];
                  if (!se2.matched) continue;
                  for (var ti3 = 0; ti3 < se2.texts.length; ti3++) {
                    var tx2 = se2.texts[ti3];
                    if (!testMatch(tx2.text)) continue;
                    var m = { slideIndex: si, shapeId: se2.shapeId, shapeName: se2.shapeName, source: tx2.source, text: tx2.text };
                    if (tx2.source === "tableCell") { m.row = tx2.row; m.col = tx2.col; }
                    slideResults.push(m);
                  }
                }
                if (noteMatched) {
                  slideResults.push({ slideIndex: si, source: "note", text: noteText });
                }
              }
            }
          }

          var result = { query: query, caseSensitive: caseSensitive, regex: useRegex, totalSlides: total };
          if (ctxLevel === "none") {
            result.matchingSlides = slideResults;
          } else {
            result.matches = slideResults;
          }
          return result;
        `
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: format_shapes ---
  server.tool(
    'format_shapes',
    'Apply formatting to multiple shapes on a slide in one call. Generates Office.js code internally. Use for fill color, font bold/italic/size/color/name. Cannot set corner radius or borders (use edit_slide_xml code mode for those).',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      shapes: z
        .array(
          z.object({
            id: z.string().describe('Shape ID from inspect_slide or scan_slide'),
            fill: z.string().optional().describe('Fill color as hex without # (e.g., "1A1A1E")'),
            font: z
              .object({
                bold: z.boolean().optional(),
                italic: z.boolean().optional(),
                size: z.number().optional().describe('Font size in points'),
                color: z.string().optional().describe('Font color as hex without # (e.g., "FFFFFF")'),
                name: z.string().optional().describe('Font name (e.g., "Calibri")'),
              })
              .optional(),
          }),
        )
        .min(1)
        .describe('Shapes to format with their properties'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideIndex, shapes, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)

        // Build Office.js code that applies formatting to each shape
        const shapeOps = shapes
          .map((s) => {
            const lines: string[] = []
            lines.push(`  var s = shapeMap["${s.id}"];`)
            lines.push(
              `  if (!s) throw new Error("Shape " + ${JSON.stringify(s.id)} + " not found on slide ${slideIndex}");`,
            )

            if (s.fill) {
              lines.push(`  s.fill.setSolidColor("${s.fill}");`)
            }

            if (s.font) {
              lines.push(`  var tf = s.getTextFrameOrNullObject();`)
              lines.push(`  await context.sync();`)
              lines.push(`  if (!tf.isNullObject) {`)
              lines.push(`    var tr = tf.textRange;`)
              if (s.font.bold !== undefined) lines.push(`    tr.font.bold = ${s.font.bold};`)
              if (s.font.italic !== undefined) lines.push(`    tr.font.italic = ${s.font.italic};`)
              if (s.font.size !== undefined) lines.push(`    tr.font.size = ${s.font.size};`)
              if (s.font.color !== undefined) lines.push(`    tr.font.color = "${s.font.color}";`)
              if (s.font.name !== undefined) lines.push(`    tr.font.name = "${s.font.name}";`)
              lines.push(`  }`)
            }

            return lines.join('\n')
          })
          .join('\n')

        const code = `
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
var slide = slides.items[${slideIndex}];
slide.shapes.load("items");
await context.sync();
var shapeMap = {};
for (var i = 0; i < slide.shapes.items.length; i++) {
  shapeMap[slide.shapes.items[i].id] = slide.shapes.items[i];
}
${shapeOps}
await context.sync();
return { success: true, shapesFormatted: ${shapes.length} };`

        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: execute_officejs ---
  server.tool(
    'execute_officejs',
    "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. For positioning, all values are in points (1 point = 1/72 inch). Common operations: add shapes, set text, change colors, add/delete slides.",
    {
      code: z
        .string()
        .describe(
          "Office.js code to execute. Runs inside PowerPoint.run() with 'context' available. Use 'return' to send back a result.",
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ code, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const result = await pool.sendCommand('executeCode', { code }, target.ws)
        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: search_fluent_icons ---
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
    async ({ query, top, style }) => {
      try {
        const results = await searchIcons(query, top ?? 10, style)
        return { content: [{ type: 'text' as const, text: JSON.stringify(results, null, 2) }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: read_speaker_notes ---
  server.tool(
    'read_speaker_notes',
    'Read speaker notes from slides. Uses full-deck export (cached by revision) to parse notes from OOXML — Office.js has no notes API. Returns plain text per slide.',
    {
      slideRange: z.string().optional().describe('Slide indices, e.g. "0", "0-5", "2,4,7". Omit to read all slides.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ slideRange, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const localPath = await getLocalCopyPath(pool, target)
        const data = readFileSync(localPath)
        const zip = await JSZip.loadAsync(data)

        const indices = parseSlideRange(slideRange)
        const notesMap = await readNotesFromDeck(zip, indices)

        const slides = [...notesMap.entries()]
          .sort(([a], [b]) => a - b)
          .map(([slideIndex, notes]) => ({ slideIndex, notes }))

        return {
          content: [{ type: 'text' as const, text: JSON.stringify({ slides }, null, 2) }],
        }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )

  // --- Tool: edit_speaker_notes ---
  server.tool(
    'edit_speaker_notes',
    'Set speaker notes on one or more slides. Accepts markdown text (paragraphs, **bold**, *italic*). Converts to OOXML server-side and injects via single-slide reimport. Replaces all existing notes on the specified slides.',
    {
      notes: z
        .array(
          z.object({
            slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
            text: z
              .string()
              .describe(
                'Notes content. Supports markdown: **bold**, *italic*, paragraphs separated by blank lines. Use empty string "" to clear notes.',
              ),
          }),
        )
        .min(1)
        .describe('Array of slide notes to set.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    async ({ notes, presentationId }) => {
      try {
        const target = pool.resolveTarget(presentationId)
        const results: Array<{ slideIndex: number; success: boolean; error?: string }> = []

        for (const entry of notes) {
          try {
            // Export the single slide
            const exported = await exportSlide(pool, entry.slideIndex, target.ws)
            const { zip } = await extractZipFiles(exported.base64)

            // Read current slide rels
            const slideRelsFile = zip.file('ppt/slides/_rels/slide1.xml.rels')
            const slideRelsXml = slideRelsFile
              ? await slideRelsFile.async('string')
              : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`

            // Build the notes injection files
            const files = buildNotesInjection(slideRelsXml, entry.text)

            // Detect new files for Content_Types
            const existingPaths = new Set(listZipPaths(zip))
            const newPaths = Object.keys(files).filter((p) => !existingPaths.has(p))

            // Apply changes
            const modifiedBase64 = await updateZipFiles(zip, files)

            // Auto-register Content_Types for new notes files
            if (newPaths.length > 0 && !files['[Content_Types].xml']) {
              const { zip: updatedZip } = await extractZipFiles(modifiedBase64)
              await autoRegisterContentTypes(updatedZip, newPaths)
              const finalBase64 = await updatedZip.generateAsync({ type: 'base64' })
              await reimportSlide(pool, finalBase64, exported.slideId, exported.prevSlideId, target.ws)
            } else {
              await reimportSlide(pool, modifiedBase64, exported.slideId, exported.prevSlideId, target.ws)
            }

            results.push({ slideIndex: entry.slideIndex, success: true })
          } catch (err: unknown) {
            const message = err instanceof Error ? err.message : String(err)
            results.push({ slideIndex: entry.slideIndex, success: false, error: message })
          }
        }

        // Invalidate local copy cache so read_speaker_notes sees updates
        localCopyCache.delete(target.presentationId)

        const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
        const text = JSON.stringify({ results }, null, 2) + (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err)
        return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true }
      }
    },
  )
}
