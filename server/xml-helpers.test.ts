import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  autoRegisterContentTypes,
  buildEditedZipBase64,
  escapeXml,
  extractDeckText,
  extractLayoutsFromZip,
  extractParagraphs,
  extractSlideXmlFromZip,
  extractZipFiles,
  findShapeById,
  listZipPaths,
  parseSlideXml,
  replaceParagraphs,
  replaceShape,
  serializeXml,
  updateSlideXmlInZip,
  updateZipFiles,
} from './xml-helpers.ts'

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

describe('xml-helpers', () => {
  describe('findShapeById', () => {
    it('finds shape by ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')
      expect(shape).not.toBeNull()
      expect(serializeXml(shape!)).toContain('Title 1')
    })

    it('finds shape with different ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '5')
      expect(shape).not.toBeNull()
      expect(serializeXml(shape!)).toContain('Content 2')
    })

    it('returns null for non-existent ID', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      expect(findShapeById(doc, '999')).toBeNull()
    })
  })

  describe('extractParagraphs', () => {
    it('extracts <a:p> elements as XML string', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')!
      const result = extractParagraphs(shape)
      expect(result).toContain('<a:p')
      expect(result).toContain('Hello')
      expect(result).toContain('World')
      expect(result).toContain('b="1"')
      // Should NOT contain bodyPr or lstStyle
      expect(result).not.toContain('<a:bodyPr')
      expect(result).not.toContain('<a:lstStyle')
    })

    it('throws for shape without txBody', () => {
      const xml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:nvSpPr><p:cNvPr id="1" name="NoText"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
      </p:sp>`
      const doc = parseSlideXml(xml)
      const shapes = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sp')
      expect(() => extractParagraphs(shapes[0]!)).toThrow('no text body')
    })
  })

  describe('replaceParagraphs', () => {
    it('replaces paragraphs while preserving bodyPr and lstStyle', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '2')!
      const newXml =
        '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Replaced</a:t></a:r></a:p>'
      replaceParagraphs(doc, shape, newXml)
      const result = serializeXml(doc)
      expect(result).toContain('Replaced')
      expect(result).not.toContain('Hello')
      expect(result).not.toContain('World')
      // bodyPr should still be present
      expect(result).toContain('anchor="ctr"')
    })
  })

  describe('replaceShape', () => {
    it('replaces entire shape element', () => {
      const doc = parseSlideXml(SAMPLE_SLIDE_XML)
      const shape = findShapeById(doc, '5')!
      const newShapeXml = `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="5" name="Replaced Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>New content</a:t></a:r></a:p></p:txBody>
      </p:sp>`
      replaceShape(doc, shape, newShapeXml)
      const result = serializeXml(doc)
      expect(result).toContain('Replaced Shape')
      expect(result).toContain('New content')
      expect(result).not.toContain('Body text')
    })
  })

  describe('zip helpers', () => {
    it('extracts and updates slide XML in zip', async () => {
      // Create a test zip with slide XML
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      const base64 = await zip.generateAsync({ type: 'base64' })

      // Extract
      const { zip: extractedZip, xmlString } = await extractSlideXmlFromZip(base64)
      expect(xmlString).toContain('Hello')
      expect(xmlString).toContain('World')

      // Modify and update
      const modified = xmlString.replace('Hello', 'Modified')
      const updatedBase64 = await updateSlideXmlInZip(extractedZip, modified)

      // Verify round-trip
      const { xmlString: finalXml } = await extractSlideXmlFromZip(updatedBase64)
      expect(finalXml).toContain('Modified')
      expect(finalXml).not.toContain('Hello')
    })

    it('throws when slide XML not found in zip', async () => {
      const zip = new JSZip()
      zip.file('ppt/other.xml', '<root/>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      await expect(extractSlideXmlFromZip(base64)).rejects.toThrow('File not found in zip')
    })
  })

  describe('multi-file zip helpers', () => {
    it('listZipPaths returns sorted paths', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', '<slide/>')
      zip.file('[Content_Types].xml', '<Types/>')
      zip.file('ppt/charts/chart1.xml', '<chart/>')
      const paths = listZipPaths(zip)
      expect(paths).toEqual(['[Content_Types].xml', 'ppt/charts/chart1.xml', 'ppt/slides/slide1.xml'])
    })

    it('extractZipFiles extracts specific paths', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('ppt/charts/chart1.xml', '<c:chartSpace/>')
      zip.file('ppt/media/image1.png', 'binarydata')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { files } = await extractZipFiles(base64, ['ppt/slides/slide1.xml', 'ppt/charts/chart1.xml'])
      expect(Object.keys(files)).toHaveLength(2)
      expect(files['ppt/slides/slide1.xml']).toContain('Hello')
      expect(files['ppt/charts/chart1.xml']).toContain('chartSpace')
    })

    it('extractZipFiles throws for missing path', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      const base64 = await zip.generateAsync({ type: 'base64' })

      await expect(extractZipFiles(base64, ['ppt/slides/slide1.xml', 'ppt/missing.xml'])).rejects.toThrow(
        'File not found in zip: ppt/missing.xml',
      )
    })

    it('extractZipFiles auto-discovers text/xml files when no paths given', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types/>')
      zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships/>')
      zip.file('ppt/media/image1.png', 'binarydata')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { files } = await extractZipFiles(base64)
      expect(files['ppt/slides/slide1.xml']).toBeDefined()
      expect(files['[Content_Types].xml']).toBeDefined()
      expect(files['ppt/slides/_rels/slide1.xml.rels']).toBeDefined()
      // Binary files should NOT be included
      expect(files['ppt/media/image1.png']).toBeUndefined()
    })

    it('updateZipFiles updates multiple files and adds new ones', async () => {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', SAMPLE_SLIDE_XML)
      zip.file('[Content_Types].xml', '<Types/>')
      const base64 = await zip.generateAsync({ type: 'base64' })

      const { zip: extractedZip } = await extractZipFiles(base64, ['ppt/slides/slide1.xml'])
      const updatedBase64 = await updateZipFiles(extractedZip, {
        'ppt/slides/slide1.xml': SAMPLE_SLIDE_XML.replace('Hello', 'Updated'),
        'ppt/charts/chart1.xml': '<c:chartSpace><c:chart/></c:chartSpace>',
      })

      // Verify round-trip
      const { files } = await extractZipFiles(updatedBase64, [
        'ppt/slides/slide1.xml',
        'ppt/charts/chart1.xml',
        '[Content_Types].xml',
      ])
      expect(files['ppt/slides/slide1.xml']).toContain('Updated')
      expect(files['ppt/charts/chart1.xml']).toContain('<c:chart/>')
      expect(files['[Content_Types].xml']).toBe('<Types/>') // untouched
    })
  })

  describe('autoRegisterContentTypes', () => {
    it('adds chart Override when missing', async () => {
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      )
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('PartName="/ppt/charts/chart1.xml"')
      expect(ct).toContain('drawingml.chart+xml')
    })

    it('does not duplicate existing Override', async () => {
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      )
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      // Should appear exactly once
      const matches = ct.match(/PartName="\/ppt\/charts\/chart1.xml"/g)
      expect(matches).toHaveLength(1)
    })

    it('ignores paths that do not match known patterns', async () => {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<Types></Types>')
      await autoRegisterContentTypes(zip, ['ppt/slides/slide1.xml', 'ppt/unknown/foo.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toBe('<Types></Types>') // unchanged
    })

    it('handles multiple chart files', async () => {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<Types></Types>')
      await autoRegisterContentTypes(zip, ['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml'])
      const ct = await zip.file('[Content_Types].xml')!.async('string')
      expect(ct).toContain('PartName="/ppt/charts/chart1.xml"')
      expect(ct).toContain('PartName="/ppt/charts/chart2.xml"')
    })
  })

  describe('escapeXml', () => {
    it('escapes all XML special characters', () => {
      expect(escapeXml('&')).toBe('&amp;')
      expect(escapeXml('<')).toBe('&lt;')
      expect(escapeXml('>')).toBe('&gt;')
      expect(escapeXml('"')).toBe('&quot;')
      expect(escapeXml("'")).toBe('&apos;')
    })

    it('escapes mixed content', () => {
      expect(escapeXml('Q1 Revenue & "Margins" <2026>')).toBe('Q1 Revenue &amp; &quot;Margins&quot; &lt;2026&gt;')
    })

    it('passes through normal text unchanged', () => {
      expect(escapeXml('Hello World 123')).toBe('Hello World 123')
    })
  })

  describe('extractLayoutsFromZip', () => {
    const MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`

    function makeLayoutXml(opts: { layoutType?: string; name?: string; shapes?: string }): string {
      const typeAttr = opts.layoutType ? ` type="${opts.layoutType}"` : ''
      const nameAttr = opts.name ? ` name="${opts.name}"` : ''
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${typeAttr}>
  <p:cSld${nameAttr}>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${opts.shapes ?? ''}
    </p:spTree>
  </p:cSld>
</p:sldLayout>`
    }

    async function buildZip(layoutXml: string): Promise<JSZip> {
      const zip = new JSZip()
      zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', MASTER_RELS)
      zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml)
      return zip
    }

    it('extracts basic placeholder type, idx, and name', async () => {
      const xml = makeLayoutXml({
        name: 'Two Content',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Content Placeholder 2"/>
              <p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts).toHaveLength(1)
      expect(layouts[0]!.name).toBe('Two Content')
      expect(layouts[0]!.placeholders).toHaveLength(2)

      const title = layouts[0]!.placeholders[0]!
      expect(title.type).toBe('title')
      expect(title.idx).toBeUndefined()
      expect(title.name).toBe('Title 1')

      const content = layouts[0]!.placeholders[1]!
      expect(content.type).toBe('obj') // default per OOXML spec
      expect(content.idx).toBe(1)
      expect(content.name).toBe('Content Placeholder 2')
    })

    it('extracts position and size from xfrm (EMU → points)', async () => {
      const xml = makeLayoutXml({
        name: 'With Position',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="838200" y="365125"/>
                <a:ext cx="10515600" cy="1325563"/>
              </a:xfrm>
            </p:spPr>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      const ph = layouts[0]!.placeholders[0]!
      // 838200 / 12700 = 66
      expect(ph.left).toBe(66)
      // 365125 / 12700 ≈ 28.75
      expect(ph.top).toBe(28.75)
      // 10515600 / 12700 ≈ 828
      expect(ph.width).toBe(828)
      // 1325563 / 12700 ≈ 104.38
      expect(ph.height).toBe(104.38)
    })

    it('leaves position undefined when xfrm is missing (inherited from master)', async () => {
      const xml = makeLayoutXml({
        name: 'Inherited',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      const ph = layouts[0]!.placeholders[0]!
      expect(ph.left).toBeUndefined()
      expect(ph.top).toBeUndefined()
      expect(ph.width).toBeUndefined()
      expect(ph.height).toBeUndefined()
    })

    it('reads layout type from sldLayout element', async () => {
      const xml = makeLayoutXml({ name: 'Two Objects', layoutType: 'twoObj' })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.type).toBe('twoObj')
    })

    it('leaves layout type undefined when not set', async () => {
      const xml = makeLayoutXml({ name: 'Custom Layout' })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.type).toBeUndefined()
    })

    it('defaults placeholder type to "obj" per OOXML spec', async () => {
      const xml = makeLayoutXml({
        name: 'Default Type',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Content 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.placeholders[0]!.type).toBe('obj')
    })

    it('reads sz attribute on placeholder', async () => {
      const xml = makeLayoutXml({
        name: 'Half Size',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Content 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph sz="half" idx="1"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.placeholders[0]!.sz).toBe('half')
    })

    it('reads description (alt text) from cNvPr', async () => {
      const xml = makeLayoutXml({
        name: 'With Description',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Hero Image" descr="16:9 landscape photo, no text overlay"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="3"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      const ph = layouts[0]!.placeholders[0]!
      expect(ph.type).toBe('pic')
      expect(ph.name).toBe('Hero Image')
      expect(ph.description).toBe('16:9 landscape photo, no text overlay')
    })

    it('filters utility placeholders (sldNum, ftr, dt, hdr)', async () => {
      const xml = makeLayoutXml({
        name: 'With Utilities',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Slide Number"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="12" sz="quarter"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Footer"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="3" sz="quarter"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Date"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="dt" idx="10"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Header"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="hdr" idx="11"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.placeholders).toHaveLength(1)
      expect(layouts[0]!.placeholders[0]!.type).toBe('title')
    })

    it('skips non-placeholder shapes', async () => {
      const xml = makeLayoutXml({
        name: 'Mixed Shapes',
        shapes: `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Decorative Rectangle"/>
              <p:cNvSpPr/><p:nvPr/>
            </p:nvSpPr>
            <p:spPr/>
          </p:sp>`,
      })
      const layouts = await extractLayoutsFromZip(await buildZip(xml))
      expect(layouts[0]!.placeholders).toHaveLength(1)
      expect(layouts[0]!.placeholders[0]!.type).toBe('title')
    })
  })

  describe('extractDeckText slide ordering', () => {
    const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    const NS_PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    const REL_TYPE_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

    function titleSlideXml(titleText: string): string {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="${NS_R}">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>${titleText}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
    }

    /**
     * Build a zip whose <p:sldIdLst> display order is slide3 first, then slide1
     * (rId3 → slide3.xml, rId1 → slide1.xml). Filename sort would yield the
     * opposite order, so this distinguishes display order from filename order.
     */
    async function buildOrderedZip(): Promise<Buffer> {
      const zip = new JSZip()
      zip.file('ppt/slides/slide1.xml', titleSlideXml('First by filename'))
      zip.file('ppt/slides/slide3.xml', titleSlideXml('First by display order'))
      zip.file(
        'ppt/presentation.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="${NS_R}">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId3"/>
    <p:sldId id="257" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`,
      )
      zip.file(
        'ppt/_rels/presentation.xml.rels',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_PKG_RELS}">
  <Relationship Id="rId1" Type="${REL_TYPE_SLIDE}" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="${REL_TYPE_SLIDE}" Target="slides/slide3.xml"/>
</Relationships>`,
      )
      return zip.generateAsync({ type: 'nodebuffer' })
    }

    it('reports slides in presentation display order, not filename order', async () => {
      const buf = await buildOrderedZip()
      const result = await extractDeckText(buf)
      expect(result).toHaveLength(2)
      expect(result[0]!.title).toBe('First by display order')
      expect(result[1]!.title).toBe('First by filename')
      expect(result[0]!.slide).toBe(0)
    })

    it('slideRange filter uses display order index', async () => {
      const buf = await buildOrderedZip()
      const result = await extractDeckText(buf, [0])
      expect(result).toHaveLength(1)
      expect(result[0]!.title).toBe('First by display order')
      expect(result[0]!.slide).toBe(0)
    })
  })
})

describe('buildEditedZipBase64', () => {
  const CT_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '</Types>'
  const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

  async function baseZipBase64(): Promise<string> {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_XML)
    zip.file('ppt/slides/slide1.xml', '<p:sld/>')
    return await zip.generateAsync({ type: 'base64' })
  }

  it('auto-registers a new part content type exactly once when files map omits Content_Types', async () => {
    const exported = await baseZipBase64()
    const { base64, newPaths } = await buildEditedZipBase64(exported, {
      'ppt/charts/chart1.xml': '<c:chartSpace/>',
    })
    expect(newPaths).toEqual(['ppt/charts/chart1.xml'])
    const out = await JSZip.loadAsync(Buffer.from(base64, 'base64'))
    const ct = await out.file('[Content_Types].xml')!.async('string')
    const occurrences = ct.split(`PartName="/ppt/charts/chart1.xml"`).length - 1
    expect(occurrences).toBe(1)
    expect(ct).toContain(CHART_CT)
  })

  it('does not double-register when files map already provides Content_Types', async () => {
    const exported = await baseZipBase64()
    const userCt = CT_XML.replace(
      '</Types>',
      `<Override PartName="/ppt/charts/chart1.xml" ContentType="${CHART_CT}"/></Types>`,
    )
    const { base64 } = await buildEditedZipBase64(exported, {
      'ppt/charts/chart1.xml': '<c:chartSpace/>',
      '[Content_Types].xml': userCt,
    })
    const out = await JSZip.loadAsync(Buffer.from(base64, 'base64'))
    const ct = await out.file('[Content_Types].xml')!.async('string')
    const occurrences = ct.split(`PartName="/ppt/charts/chart1.xml"`).length - 1
    expect(occurrences).toBe(1)
  })
})
