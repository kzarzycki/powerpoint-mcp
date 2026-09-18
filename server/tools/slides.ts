import { readFileSync } from 'node:fs'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import JSZip from 'jszip'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import { exportSlide, extractLayoutsFromZip, extractSlideXmlFromZip, NS_P, parseSlideXml } from '../xml-helpers.ts'
import { buildInsertOptions, getConcurrentWarning, getLocalCopyPath } from './shared.ts'

export function registerSlideTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
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
    withTool(async ({ layoutName, position, placeholders, presentationId }) => {
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

      // Step 1: Read all layouts (every master) from OOXML, and confirm the
      // requested name exists before round-tripping to Office.js.
      const localPath = await getLocalCopyPath(pool, target)
      const fileData = readFileSync(localPath)
      const zip = await JSZip.loadAsync(fileData)
      const layouts = await extractLayoutsFromZip(zip)
      const targetLower = layoutName.toLowerCase()
      if (!layouts.some((l) => l.name.toLowerCase() === targetLower)) {
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

      // Step 2: Add slide via Office.js (find layout, add, moveTo). Multiple
      // masters can share a layout name (#120), so the search here — not a
      // second independent OOXML name search — is the single source of
      // truth for which layout was actually inserted. It reports back the
      // matched master/layout position so step 3 can look up the *same*
      // layout's placeholder metadata by identity, not by re-matching name.
      const addCode = `
          var masters = context.presentation.slideMasters;
          masters.load("items");
          await context.sync();
          for (var m = 0; m < masters.items.length; m++) {
            masters.items[m].layouts.load("items/id,items/name");
          }
          await context.sync();
          var targetName = ${JSON.stringify(layoutName)}.toLowerCase();
          var layout = null;
          var matchedMasterIndex = -1;
          var matchedLayoutIndex = -1;
          for (var m = 0; m < masters.items.length && !layout; m++) {
            var ml = masters.items[m].layouts.items;
            for (var i = 0; i < ml.length; i++) {
              if (ml[i].name.toLowerCase() === targetName) {
                layout = ml[i];
                matchedMasterIndex = m;
                matchedLayoutIndex = i;
                break;
              }
            }
          }
          if (!layout) {
            var names = [];
            for (var m = 0; m < masters.items.length; m++) {
              var ml2 = masters.items[m].layouts.items;
              for (var i = 0; i < ml2.length; i++) { names.push(ml2[i].name); }
            }
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
          return { slideIndex: finalIndex, slideId: newSlide.id, slideCount: slides.items.length, layoutName: layout.name, masterIndex: matchedMasterIndex, layoutIndexInMaster: matchedLayoutIndex };
        `
      const addResult = (await pool.sendCommand('executeCode', { code: addCode }, target.ws)) as {
        slideIndex: number
        slideId: string
        slideCount: number
        layoutName: string
        masterIndex: number
        layoutIndexInMaster: number
      }

      // Resolve the exact OOXML layout Office.js selected, by identity — not
      // by re-searching by name, which is what let #120 silently mix
      // placeholder metadata from the wrong master. Name-match is kept only
      // as a defensive fallback for a malformed/mismatched file.
      const layoutInfo =
        layouts.find(
          (l) => l.masterIndex === addResult.masterIndex && l.layoutIndexInMaster === addResult.layoutIndexInMaster,
        ) ?? layouts.find((l) => l.name.toLowerCase() === targetLower)!

      // Build idx→name map from the matched layout's placeholders
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
    }),
  )

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
    withTool(async ({ slideIndex, insertAfter, presentationId }) => {
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
    }),
  )

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
    withTool(
      async ({ sourceSlideIndex, sourcePresentationId, destinationPresentationId, targetSlideId, formatting }) => {
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
        const optionsArg = buildInsertOptions(formatting, targetSlideId)

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
      },
    ),
  )
}
