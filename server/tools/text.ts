import { readFileSync } from 'node:fs'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import {
  exportSlide,
  extractDeckText,
  extractParagraphs,
  extractSlideXmlFromZip,
  findShapeById,
  parseSlideXml,
  reimportSlide,
  replaceParagraphs,
  serializeXml,
  updateSlideXmlInZip,
} from '../xml-helpers.ts'
import { buildFormatShapeOps, getConcurrentWarning, getLocalCopyPath, parseSlideRange } from './shared.ts'

export function registerTextTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
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
    withTool(async ({ slideIndex, shapeId, xml, presentationId }) => {
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
    }),
  )

  server.tool(
    'read_shape_paragraphs',
    "Read raw OOXML <a:p> paragraphs from a shape's text body. Returns the paragraph XML as a string — preserves all formatting (bold, colors, bullets, etc.) that textRange.text strips. Use with the /pptx skill's OOXML knowledge to understand and modify the XML.",
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index from scan_slide results'),
      shapeId: z.string().describe('Shape ID from inspect_slide results (e.g. "5")'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideIndex, shapeId, presentationId }) => {
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
    }),
  )

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
    withTool(async ({ slideRange, includeNotes, presentationId }) => {
      const target = pool.resolveTarget(presentationId)
      const localPath = await getLocalCopyPath(pool, target)
      const zipBuffer = readFileSync(localPath)
      const indices = slideRange ? parseSlideRange(slideRange) : null
      const result = await extractDeckText(zipBuffer, indices, includeNotes === true)
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(result) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

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
    withTool(
      async ({ query, slideRange, caseSensitive, regex, context: contextLevel, includeNotes, presentationId }) => {
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
      },
    ),
  )

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
    withTool(async ({ slideIndex, shapes, presentationId }) => {
      const target = pool.resolveTarget(presentationId)

      // Build Office.js code that applies formatting to each shape
      const shapeOps = buildFormatShapeOps(shapes, slideIndex)

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
    }),
  )
}
