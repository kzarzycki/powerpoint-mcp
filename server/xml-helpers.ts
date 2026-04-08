import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import JSZip from 'jszip'
import type { WebSocket } from 'ws'
import type { ConnectionPool } from './bridge.ts'

const SLIDE_XML_PATH = 'ppt/slides/slide1.xml'

// ---------------------------------------------------------------------------
// Export slide as base64 zip via add-in
// ---------------------------------------------------------------------------

export interface ExportedSlide {
  base64: string
  slideId: string
  prevSlideId: string | null
}

export async function exportSlide(
  pool: ConnectionPool,
  slideIndex: number,
  targetWs: WebSocket,
  timeout?: number,
): Promise<ExportedSlide> {
  const code = `
    var slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    if (${slideIndex} >= slides.items.length) {
      throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
    }
    var slide = slides.items[${slideIndex}];
    var result = slide.exportAsBase64();
    await context.sync();
    var prevSlideId = ${slideIndex} > 0 ? slides.items[${slideIndex} - 1].id : null;
    return { base64: result.value, slideId: slide.id, prevSlideId: prevSlideId };
  `
  return (await pool.sendCommand('executeCode', { code }, targetWs, timeout)) as ExportedSlide
}

// ---------------------------------------------------------------------------
// Reimport modified slide (delete old + insert modified at same position)
// ---------------------------------------------------------------------------

export async function reimportSlide(
  pool: ConnectionPool,
  modifiedBase64: string,
  slideId: string,
  prevSlideId: string | null,
  targetWs: WebSocket,
  timeout?: number,
): Promise<void> {
  const optionsParts = ['formatting: "KeepSourceFormatting"']
  if (prevSlideId) {
    optionsParts.push(`targetSlideId: "${prevSlideId}"`)
  }
  const optionsStr = `{ ${optionsParts.join(', ')} }`

  // The base64 is embedded directly in the code string — it stays server-side
  // and never enters Claude's context.
  const code = `
    var slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    var countBefore = slides.items.length;
    // Find and delete the original slide by ID
    var found = false;
    for (var i = 0; i < slides.items.length; i++) {
      if (slides.items[i].id === "${slideId}") {
        slides.items[i].delete();
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error("Original slide not found for reimport (ID: ${slideId})");
    }
    // Batch delete + insert in one sync to reduce the window for partial failure
    context.presentation.insertSlidesFromBase64("${modifiedBase64}", ${optionsStr});
    await context.sync();
    // Verify slide count is unchanged (deleted one, inserted one)
    slides.load("items");
    await context.sync();
    if (slides.items.length !== countBefore) {
      throw new Error("Reimport verification failed: expected " + countBefore + " slides but found " + slides.items.length);
    }
    return { success: true };
  `
  await pool.sendCommand('executeCode', { code }, targetWs, timeout)
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

export function parseSlideXml(xmlString: string): Document {
  return new DOMParser().parseFromString(xmlString, 'text/xml')
}

export function serializeXml(doc: Document | Element): string {
  return new XMLSerializer().serializeToString(doc)
}

/**
 * Find a <p:sp> shape element by matching <p:cNvPr id="shapeId">.
 * Shape IDs from Office.js are strings like "5" — the OOXML id attribute matches.
 */
export function findShapeById(doc: Document, shapeId: string): Element | null {
  const shapes = doc.getElementsByTagNameNS(NS_P, 'sp')
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i]!
    const nvSpPr = shape.getElementsByTagNameNS(NS_P, 'nvSpPr')
    if (nvSpPr.length === 0) continue
    const cNvPr = nvSpPr[0]!.getElementsByTagNameNS(NS_P, 'cNvPr')
    if (cNvPr.length > 0 && cNvPr[0]!.getAttribute('id') === shapeId) {
      return shape
    }
  }
  return null
}

/**
 * Extract all <a:p> elements from a shape's <p:txBody> as serialized XML string.
 */
export function extractParagraphs(shape: Element): string {
  const txBody = shape.getElementsByTagNameNS(NS_P, 'txBody')
  if (txBody.length === 0) {
    throw new Error('Shape has no text body (<p:txBody>)')
  }
  const body = txBody[0]!
  const paragraphs = body.getElementsByTagNameNS(NS_A, 'p')
  const parts: string[] = []
  for (let i = 0; i < paragraphs.length; i++) {
    parts.push(serializeXml(paragraphs[i]!))
  }
  return parts.join('')
}

/**
 * Replace all <a:p> elements in a shape's <p:txBody>, preserving <a:bodyPr> and <a:lstStyle>.
 */
export function replaceParagraphs(doc: Document, shape: Element, paragraphXml: string): void {
  const txBody = shape.getElementsByTagNameNS(NS_P, 'txBody')
  if (txBody.length === 0) {
    throw new Error('Shape has no text body (<p:txBody>)')
  }
  const body = txBody[0]!

  // Save <a:bodyPr> and <a:lstStyle>
  const bodyPr = body.getElementsByTagNameNS(NS_A, 'bodyPr')[0] ?? null
  const lstStyle = body.getElementsByTagNameNS(NS_A, 'lstStyle')[0] ?? null

  // Remove all children
  while (body.firstChild) {
    body.removeChild(body.firstChild)
  }

  // Re-append preserved elements
  if (bodyPr) body.appendChild(bodyPr)
  if (lstStyle) body.appendChild(lstStyle)

  // Parse and import new paragraphs
  const wrapper = `<wrapper xmlns:a="${NS_A}">${paragraphXml}</wrapper>`
  const fragDoc = new DOMParser().parseFromString(wrapper, 'text/xml')
  const newParagraphs = fragDoc.documentElement.childNodes
  for (let i = 0; i < newParagraphs.length; i++) {
    const imported = doc.importNode(newParagraphs[i]!, true)
    body.appendChild(imported)
  }
}

/**
 * Replace a shape's <p:sp> element in the document with new XML.
 */
export function replaceShape(doc: Document, oldShape: Element, newShapeXml: string): void {
  const fragDoc = new DOMParser().parseFromString(newShapeXml, 'text/xml')
  const imported = doc.importNode(fragDoc.documentElement, true)
  oldShape.parentNode!.replaceChild(imported, oldShape)
}

// ---------------------------------------------------------------------------
// Zip helpers
// ---------------------------------------------------------------------------

/** List all file paths in a zip (excludes directory entries). */
export function listZipPaths(zip: JSZip): string[] {
  const paths: string[] = []
  zip.forEach((relativePath, file) => {
    if (!file.dir) paths.push(relativePath)
  })
  return paths.sort()
}

/** Extract specific files from a base64 zip. If paths omitted, extracts all text/xml files. */
export async function extractZipFiles(
  base64: string,
  paths?: string[],
): Promise<{ zip: JSZip; files: Record<string, string> }> {
  const zip = await JSZip.loadAsync(base64, { base64: true })
  const files: Record<string, string> = {}

  if (paths) {
    for (const path of paths) {
      const file = zip.file(path)
      if (!file) {
        throw new Error(`File not found in zip: ${path}`)
      }
      files[path] = await file.async('string')
    }
  } else {
    // Auto-discover: extract all text/xml files (skip binary media)
    const allPaths = listZipPaths(zip)
    for (const path of allPaths) {
      if (path.endsWith('/')) continue // skip directories
      if (path.match(/\.(xml|rels)$/) || path === '[Content_Types].xml') {
        const file = zip.file(path)
        if (file) {
          files[path] = await file.async('string')
        }
      }
    }
  }

  return { zip, files }
}

/** Update multiple files in a zip and regenerate as base64. Can add new files. */
export async function updateZipFiles(zip: JSZip, files: Record<string, string>): Promise<string> {
  for (const [path, content] of Object.entries(files)) {
    zip.file(path, content)
  }
  return await zip.generateAsync({ type: 'base64' })
}

const CONTENT_TYPE_MAP: Record<string, string> = {
  'ppt/charts/': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  'ppt/notesSlides/': 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
}

/**
 * Auto-register Content_Types for newly added files.
 * Reads [Content_Types].xml, adds <Override> entries for known file types,
 * and writes back. Only adds entries not already present.
 */
export async function autoRegisterContentTypes(zip: JSZip, newPaths: string[]): Promise<void> {
  const overrides: Array<{ partName: string; contentType: string }> = []
  for (const path of newPaths) {
    for (const [prefix, contentType] of Object.entries(CONTENT_TYPE_MAP)) {
      if (path.startsWith(prefix) && path.endsWith('.xml')) {
        overrides.push({ partName: `/${path}`, contentType })
      }
    }
  }
  if (overrides.length === 0) return

  const ctFile = zip.file('[Content_Types].xml')
  if (!ctFile) return

  let ctXml = await ctFile.async('string')
  for (const { partName, contentType } of overrides) {
    if (ctXml.includes(`PartName="${partName}"`)) continue
    // Insert before closing </Types>
    const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`
    ctXml = ctXml.replace('</Types>', `${override}</Types>`)
  }
  zip.file('[Content_Types].xml', ctXml)
}

// Legacy wrappers — used by existing read/edit_slide_text and read/edit_slide_xml tools
export async function extractSlideXmlFromZip(base64: string): Promise<{ zip: JSZip; xmlString: string }> {
  const { zip, files } = await extractZipFiles(base64, [SLIDE_XML_PATH])
  return { zip, xmlString: files[SLIDE_XML_PATH]! }
}

export async function updateSlideXmlInZip(zip: JSZip, xmlString: string): Promise<string> {
  return await updateZipFiles(zip, { [SLIDE_XML_PATH]: xmlString })
}

// ---------------------------------------------------------------------------
// Theme extraction from exported slide zip
// ---------------------------------------------------------------------------

export interface ThemeInfo {
  name: string
  colors: Record<string, string>
  fonts: { major: string; minor: string }
}

export async function extractThemeFromZip(base64: string): Promise<ThemeInfo> {
  const zip = await JSZip.loadAsync(base64, { base64: true })
  // Find the theme file (usually ppt/theme/theme1.xml)
  const themePath = Object.keys(zip.files).find((p) => p.startsWith('ppt/theme/') && p.endsWith('.xml'))
  if (!themePath) throw new Error('No theme file found in zip')
  const themeXml = await zip.file(themePath)!.async('string')
  const doc = new DOMParser().parseFromString(themeXml, 'text/xml')

  // Extract color scheme
  const clrScheme = doc.getElementsByTagNameNS(NS_A, 'clrScheme')[0]
  const colors: Record<string, string> = {}
  if (clrScheme) {
    for (let i = 0; i < clrScheme.childNodes.length; i++) {
      const node = clrScheme.childNodes[i] as Element
      if (node.nodeType !== 1) continue // skip text nodes
      const tag = node.localName
      // Color value is in the first child element's val attribute (srgbClr or sysClr)
      const valElem = node.getElementsByTagNameNS(NS_A, 'srgbClr')[0] ?? node.getElementsByTagNameNS(NS_A, 'sysClr')[0]
      if (valElem) {
        colors[tag] = valElem.getAttribute('val') ?? valElem.getAttribute('lastClr') ?? ''
      }
    }
  }

  // Extract font scheme
  const fontScheme = doc.getElementsByTagNameNS(NS_A, 'fontScheme')[0]
  const majorLatin = fontScheme?.getElementsByTagNameNS(NS_A, 'majorFont')[0]?.getElementsByTagNameNS(NS_A, 'latin')[0]
  const minorLatin = fontScheme?.getElementsByTagNameNS(NS_A, 'minorFont')[0]?.getElementsByTagNameNS(NS_A, 'latin')[0]

  return {
    name: clrScheme?.getAttribute('name') ?? 'Unknown',
    colors,
    fonts: {
      major: majorLatin?.getAttribute('typeface') ?? '',
      minor: minorLatin?.getAttribute('typeface') ?? '',
    },
  }
}

// ---------------------------------------------------------------------------
// Layout extraction from full presentation zip
// ---------------------------------------------------------------------------

const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const LAYOUT_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

// TODO: consider reading <p:tag> elements for agent-specific metadata (agent:* prefix)

const EMU_PER_PT = 12700
function emuToPoints(emu: number): number {
  return Math.round((emu / EMU_PER_PT) * 100) / 100
}

export interface PlaceholderInfo {
  type: string // <p:ph type="...">, default "obj" per OOXML spec
  idx?: number // <p:ph idx="...">
  name?: string // <p:cNvPr name="..."> — shape name, visible in Selection Pane
  description?: string // <p:cNvPr descr="..."> — alt text, visible as tooltip
  sz?: string // <p:ph sz="..."> — "full"/"half"/"quarter"
  left?: number // points, from <a:xfrm>; undefined = inherited from master
  top?: number
  width?: number
  height?: number
}

export interface LayoutInfo {
  index: number
  name: string
  type?: string // <p:sldLayout type="...">, e.g. "blank", "twoObj", "secHead"
  placeholders: PlaceholderInfo[]
  usedBySlides?: number[]
}

export async function extractLayoutsFromZip(zip: JSZip): Promise<LayoutInfo[]> {
  // Find the first slide master's rels to get layout order
  const masterRelsPath = 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
  const masterRelsFile = zip.file(masterRelsPath)
  if (!masterRelsFile) throw new Error('No slide master rels found')
  const masterRelsXml = await masterRelsFile.async('string')
  const relsDoc = new DOMParser().parseFromString(masterRelsXml, 'text/xml')

  // Collect layout targets in document order (this determines layoutIndex)
  const layoutTargets: string[] = []
  const rels = relsDoc.getElementsByTagNameNS(NS_RELS, 'Relationship')
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i]!
    if (rel.getAttribute('Type') === LAYOUT_TYPE) {
      const target = rel.getAttribute('Target') ?? ''
      // Target is relative like "../slideLayouts/slideLayout1.xml"
      const resolved = target.replace('..', 'ppt')
      layoutTargets.push(resolved)
    }
  }

  // Parse each layout XML for name, type, and placeholders
  const layouts: LayoutInfo[] = []
  for (let i = 0; i < layoutTargets.length; i++) {
    const layoutFile = zip.file(layoutTargets[i]!)
    if (!layoutFile) continue
    const layoutXml = await layoutFile.async('string')
    const doc = new DOMParser().parseFromString(layoutXml, 'text/xml')

    // Layout type from <p:sldLayout type="...">
    const sldLayout = doc.getElementsByTagNameNS(NS_P, 'sldLayout')[0]
    const layoutType = sldLayout?.getAttribute('type') ?? undefined

    // Name from <p:cSld name="...">
    const cSld = doc.getElementsByTagNameNS(NS_P, 'cSld')[0]
    const name = cSld?.getAttribute('name') ?? `Layout ${i}`

    // Iterate shapes top-down, extract placeholder metadata
    const placeholders: PlaceholderInfo[] = []
    const shapes = doc.getElementsByTagNameNS(NS_P, 'sp')
    for (let j = 0; j < shapes.length; j++) {
      const shape = shapes[j]!
      const nvSpPr = shape.getElementsByTagNameNS(NS_P, 'nvSpPr')[0]
      if (!nvSpPr) continue
      const nvPr = nvSpPr.getElementsByTagNameNS(NS_P, 'nvPr')[0]
      if (!nvPr) continue
      const ph = nvPr.getElementsByTagNameNS(NS_P, 'ph')[0]
      if (!ph) continue // not a placeholder shape

      const phType = ph.getAttribute('type') || 'obj'

      // Skip utility placeholders — auto-filled by PowerPoint, not agent-relevant
      if (phType === 'sldNum' || phType === 'ftr' || phType === 'dt' || phType === 'hdr') continue

      const info: PlaceholderInfo = { type: phType }

      const idxStr = ph.getAttribute('idx')
      if (idxStr) info.idx = Number.parseInt(idxStr, 10)
      const szStr = ph.getAttribute('sz')
      if (szStr) info.sz = szStr

      // Shape name and description from <p:cNvPr>
      const cNvPr = nvSpPr.getElementsByTagNameNS(NS_P, 'cNvPr')[0]
      if (cNvPr) {
        const shapeName = cNvPr.getAttribute('name')
        if (shapeName) info.name = shapeName
        const descr = cNvPr.getAttribute('descr')
        if (descr) info.description = descr
      }

      // Position/size from <p:spPr><a:xfrm>
      const spPr = shape.getElementsByTagNameNS(NS_P, 'spPr')[0]
      const xfrm = spPr?.getElementsByTagNameNS(NS_A, 'xfrm')[0]
      if (xfrm) {
        const off = xfrm.getElementsByTagNameNS(NS_A, 'off')[0]
        const ext = xfrm.getElementsByTagNameNS(NS_A, 'ext')[0]
        if (off) {
          const x = off.getAttribute('x')
          const y = off.getAttribute('y')
          if (x) info.left = emuToPoints(Number.parseInt(x, 10))
          if (y) info.top = emuToPoints(Number.parseInt(y, 10))
        }
        if (ext) {
          const cx = ext.getAttribute('cx')
          const cy = ext.getAttribute('cy')
          if (cx) info.width = emuToPoints(Number.parseInt(cx, 10))
          if (cy) info.height = emuToPoints(Number.parseInt(cy, 10))
        }
      }

      placeholders.push(info)
    }

    layouts.push({ index: i, name, ...(layoutType ? { type: layoutType } : {}), placeholders })
  }

  return layouts
}
