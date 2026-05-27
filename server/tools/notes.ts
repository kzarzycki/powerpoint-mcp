import { readFileSync } from 'node:fs'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import JSZip from 'jszip'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { buildNotesInjection, readNotesFromDeck } from '../notes-helpers.ts'
import { errorMessage, withTool } from '../tool-helpers.ts'
import { applyZipEditAndReimport, exportSlide, extractZipFiles } from '../xml-helpers.ts'
import { getConcurrentWarning, getLocalCopyPath, localCopyCache, parseSlideRange } from './shared.ts'

export function registerNotesTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
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
    withTool(async ({ notes, presentationId }) => {
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

          await applyZipEditAndReimport(pool, exported, files, target.ws)

          results.push({ slideIndex: entry.slideIndex, success: true })
        } catch (err: unknown) {
          results.push({ slideIndex: entry.slideIndex, success: false, error: errorMessage(err) })
        }
      }

      // Invalidate local copy cache so read_speaker_notes sees updates
      localCopyCache.delete(target.presentationId)

      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify({ results }, null, 2) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

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
    withTool(async ({ slideRange, presentationId }) => {
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
    }),
  )
}
