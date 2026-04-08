/**
 * Speaker notes OOXML helpers.
 *
 * Office.js has no speaker notes API (tested 1.1–1.8 + preview).
 * All notes access goes through OOXML manipulation:
 *  - Reading: full-deck export → parse notesSlide XML
 *  - Writing: single-slide export → inject notesSlide XML → reimport
 */

import { DOMParser } from '@xmldom/xmldom'
import type JSZip from 'jszip'

// ── OOXML constants ────────────────────────────────────────────

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'

const REL_TYPE_NOTES_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const REL_TYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
export const NOTES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'

// ── Markdown → OOXML conversion ────────────────────────────────

/** Escape XML special characters */
export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

interface TextRun {
  text: string
  bold?: boolean
  italic?: boolean
}

/** Parse inline markdown formatting into text runs */
function parseInlineFormatting(line: string): TextRun[] {
  const runs: TextRun[] = []
  // Match bold+italic, bold, italic, or plain text
  const pattern = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|([^*]+))/g
  for (const match of line.matchAll(pattern)) {
    if (match[2]) {
      runs.push({ text: match[2], bold: true, italic: true })
    } else if (match[3]) {
      runs.push({ text: match[3], bold: true })
    } else if (match[4]) {
      runs.push({ text: match[4], italic: true })
    } else if (match[5]) {
      runs.push({ text: match[5] })
    }
  }
  return runs.length > 0 ? runs : [{ text: line }]
}

/** Convert a single text run to OOXML <a:r> element */
function runToXml(run: TextRun): string {
  const attrs: string[] = ['lang="en-US"', 'dirty="0"']
  if (run.bold) attrs.push('b="1"')
  if (run.italic) attrs.push('i="1"')
  return `<a:r><a:rPr ${attrs.join(' ')}/><a:t>${escapeXml(run.text)}</a:t></a:r>`
}

/**
 * Convert markdown text to OOXML paragraph elements.
 * Supports: paragraphs (blank line separated), **bold**, *italic*, ***bold italic***.
 * Headings (# ) are converted to bold paragraphs.
 */
export function markdownToNotesXml(text: string): string {
  if (!text || !text.trim()) {
    return '<a:p><a:endParaRPr lang="en-US"/></a:p>'
  }

  const lines = text.split('\n')
  const paragraphs: string[] = []
  let currentPara: string[] = []

  function flushParagraph() {
    if (currentPara.length === 0) return
    const combined = currentPara.join(' ')
    const runs = parseInlineFormatting(combined)
    paragraphs.push(`<a:p>${runs.map(runToXml).join('')}</a:p>`)
    currentPara = []
  }

  for (const rawLine of lines) {
    const line = rawLine.trimEnd()

    // Blank line = paragraph break
    if (!line.trim()) {
      flushParagraph()
      continue
    }

    // Heading → bold paragraph
    const headingMatch = line.match(/^#{1,6}\s+(.+)$/)
    if (headingMatch) {
      flushParagraph()
      const runs: TextRun[] = [{ text: headingMatch[1], bold: true }]
      paragraphs.push(`<a:p>${runs.map(runToXml).join('')}</a:p>`)
      continue
    }

    // Bullet → keep as-is with dash prefix
    const bulletMatch = line.match(/^[\s]*[-*+]\s+(.+)$/)
    if (bulletMatch) {
      flushParagraph()
      const runs = parseInlineFormatting(`- ${bulletMatch[1]}`)
      paragraphs.push(`<a:p>${runs.map(runToXml).join('')}</a:p>`)
      continue
    }

    // Regular line — accumulate into current paragraph
    currentPara.push(line)
  }
  flushParagraph()

  return paragraphs.length > 0 ? paragraphs.join('') : '<a:p><a:endParaRPr lang="en-US"/></a:p>'
}

// ── Notes XML extraction ───────────────────────────────────────

/** Extract plain text from a notesSlide XML string */
export function extractNotesText(notesXml: string): string | null {
  const doc = new DOMParser().parseFromString(notesXml, 'text/xml')

  // Find the body placeholder shape (type="body" idx="1")
  const shapes = doc.getElementsByTagNameNS(NS_P, 'sp')
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i]
    const phElements = shape.getElementsByTagNameNS(NS_P, 'ph')
    for (let j = 0; j < phElements.length; j++) {
      const phType = phElements[j].getAttribute('type')
      if (phType === 'body') {
        // Extract text per paragraph, joining runs within each <a:p>
        const paragraphs = shape.getElementsByTagNameNS(NS_A, 'p')
        const paraTexts: string[] = []
        for (let k = 0; k < paragraphs.length; k++) {
          const runs = paragraphs[k].getElementsByTagNameNS(NS_A, 't')
          let paraText = ''
          for (let r = 0; r < runs.length; r++) {
            paraText += runs[r].textContent || ''
          }
          if (paraText) paraTexts.push(paraText)
        }
        const result = paraTexts.join('\n')
        return result.trim() || null
      }
    }
  }
  return null
}

// ── Notes slide XML builders ───────────────────────────────────

/** Build a complete notesSlide XML with the given paragraph content */
export function buildNotesSlideXml(paragraphXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphXml}</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`
}

/** Build rels XML for a notes slide, pointing back to the slide */
export function buildNotesSlideRels(slideFileName: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_RELS}"><Relationship Id="rId1" Type="${REL_TYPE_SLIDE}" Target="../slides/${slideFileName}"/></Relationships>`
}

// ── Slide-to-notes mapping (for full-deck reads) ──────────────

interface SlideNotesMapping {
  slidePath: string
  notesPath: string | null
}

/**
 * Parse a full-deck zip to map slide indices to their notes file paths.
 * Returns a Map<slideIndex, { slidePath, notesPath }>.
 */
export async function resolveSlideToNotesMapping(zip: JSZip): Promise<Map<number, SlideNotesMapping>> {
  const mapping = new Map<number, SlideNotesMapping>()

  // Parse presentation.xml to get slide ordering
  const presFile = zip.file('ppt/presentation.xml')
  if (!presFile) return mapping
  const presXml = await presFile.async('string')
  const presDoc = new DOMParser().parseFromString(presXml, 'text/xml')

  // Parse presentation.xml.rels to resolve rIds to slide file paths
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels')
  if (!presRelsFile) return mapping
  const presRelsXml = await presRelsFile.async('string')
  const presRelsDoc = new DOMParser().parseFromString(presRelsXml, 'text/xml')

  // Build rId → target map
  const rIdToTarget = new Map<string, string>()
  const rels = presRelsDoc.getElementsByTagNameNS(NS_RELS, 'Relationship')
  for (let i = 0; i < rels.length; i++) {
    const id = rels[i].getAttribute('Id')
    const target = rels[i].getAttribute('Target')
    if (id && target) rIdToTarget.set(id, target)
  }

  // Get ordered slide list from <p:sldIdLst>
  const sldIdLst = presDoc.getElementsByTagNameNS(NS_P, 'sldId')
  for (let idx = 0; idx < sldIdLst.length; idx++) {
    const rId = sldIdLst[idx].getAttributeNS(NS_R, 'id')
    if (!rId) continue
    const target = rIdToTarget.get(rId)
    if (!target) continue

    const slidePath = target.startsWith('ppt/') ? target : `ppt/${target}`

    // Check slide's rels for a notesSlide relationship
    const slideRelsPath = `${slidePath.replace('ppt/slides/', 'ppt/slides/_rels/')}.rels`
    const slideRelsFile = zip.file(slideRelsPath)
    let notesPath: string | null = null

    if (slideRelsFile) {
      const slideRelsXml = await slideRelsFile.async('string')
      const slideRelsDoc = new DOMParser().parseFromString(slideRelsXml, 'text/xml')
      const slideRels = slideRelsDoc.getElementsByTagNameNS(NS_RELS, 'Relationship')
      for (let j = 0; j < slideRels.length; j++) {
        if (slideRels[j].getAttribute('Type') === REL_TYPE_NOTES_SLIDE) {
          const notesTarget = slideRels[j].getAttribute('Target')
          if (notesTarget) {
            notesPath = notesTarget.startsWith('ppt/') ? notesTarget : `ppt/notesSlides/${notesTarget.split('/').pop()}`
          }
          break
        }
      }
    }

    mapping.set(idx, { slidePath, notesPath })
  }

  return mapping
}

/**
 * Read notes text from a full-deck zip for the specified slide indices.
 * If slideIndices is null, reads all slides.
 */
export async function readNotesFromDeck(
  zip: JSZip,
  slideIndices: number[] | null,
): Promise<Map<number, string | null>> {
  const mapping = await resolveSlideToNotesMapping(zip)
  const result = new Map<number, string | null>()

  const indices = slideIndices ?? [...mapping.keys()]
  for (const idx of indices) {
    const entry = mapping.get(idx)
    if (!entry) {
      result.set(idx, null)
      continue
    }
    if (!entry.notesPath) {
      result.set(idx, null)
      continue
    }
    const notesFile = zip.file(entry.notesPath)
    if (!notesFile) {
      result.set(idx, null)
      continue
    }
    const notesXml = await notesFile.async('string')
    result.set(idx, extractNotesText(notesXml))
  }

  return result
}

// ── Helpers for injecting notes into a single-slide zip ────────

/**
 * Inject or replace speaker notes in a single-slide zip.
 * Returns a map of { zipPath: content } to pass to updateZipFiles.
 */
export function buildNotesInjection(slideRelsXml: string, markdownText: string): Record<string, string> {
  const files: Record<string, string> = {}
  const paragraphXml = markdownToNotesXml(markdownText)

  // The single-slide export always has slide1.xml
  files['ppt/notesSlides/notesSlide1.xml'] = buildNotesSlideXml(paragraphXml)
  files['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = buildNotesSlideRels('slide1.xml')

  // Add notesSlide relationship to slide rels if not already present
  if (!slideRelsXml.includes(REL_TYPE_NOTES_SLIDE)) {
    // Find the next available rId
    const rIdMatches = [...slideRelsXml.matchAll(/Id="rId(\d+)"/g)]
    const maxId = rIdMatches.reduce((max, m) => Math.max(max, Number(m[1])), 0)
    const newRId = `rId${maxId + 1}`
    const newRel = `<Relationship Id="${newRId}" Type="${REL_TYPE_NOTES_SLIDE}" Target="../notesSlides/notesSlide1.xml"/>`
    files['ppt/slides/_rels/slide1.xml.rels'] = slideRelsXml.replace('</Relationships>', `${newRel}</Relationships>`)
  }

  return files
}
