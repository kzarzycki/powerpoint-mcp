import { describe, expect, it } from 'vitest'
import {
  buildNotesInjection,
  buildNotesSlideRels,
  buildNotesSlideXml,
  escapeXml,
  extractNotesText,
  markdownToNotesXml,
} from './notes-helpers.ts'

describe('notes-helpers', () => {
  describe('escapeXml', () => {
    it('escapes all XML special characters', () => {
      expect(escapeXml('a & b < c > d "e" \'f\'')).toBe('a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')
    })

    it('returns plain text unchanged', () => {
      expect(escapeXml('hello world')).toBe('hello world')
    })
  })

  describe('markdownToNotesXml', () => {
    it('converts plain text to a single paragraph', () => {
      const result = markdownToNotesXml('Hello world')
      expect(result).toContain('<a:t>Hello world</a:t>')
      expect(result).toMatch(/^<a:p>.*<\/a:p>$/)
    })

    it('splits on blank lines into multiple paragraphs', () => {
      const result = markdownToNotesXml('First paragraph\n\nSecond paragraph')
      expect(result).toContain('<a:t>First paragraph</a:t>')
      expect(result).toContain('<a:t>Second paragraph</a:t>')
      const count = (result.match(/<a:p>/g) || []).length
      expect(count).toBe(2)
    })

    it('converts **bold** to b="1"', () => {
      const result = markdownToNotesXml('This is **bold** text')
      expect(result).toContain('b="1"')
      expect(result).toContain('<a:t>bold</a:t>')
    })

    it('converts *italic* to i="1"', () => {
      const result = markdownToNotesXml('This is *italic* text')
      expect(result).toContain('i="1"')
      expect(result).toContain('<a:t>italic</a:t>')
    })

    it('converts ***bold italic*** to both b="1" i="1"', () => {
      const result = markdownToNotesXml('This is ***bold italic*** text')
      expect(result).toContain('b="1" i="1"')
      expect(result).toContain('<a:t>bold italic</a:t>')
    })

    it('converts headings to bold paragraphs', () => {
      const result = markdownToNotesXml('## Section Title')
      expect(result).toContain('b="1"')
      expect(result).toContain('<a:t>Section Title</a:t>')
    })

    it('handles bullets', () => {
      const result = markdownToNotesXml('- First item\n- Second item')
      expect(result).toContain('<a:t>- First item</a:t>')
      expect(result).toContain('<a:t>- Second item</a:t>')
    })

    it('escapes XML special characters in content', () => {
      const result = markdownToNotesXml('ROI > 200% & costs < budget')
      expect(result).toContain('&gt;')
      expect(result).toContain('&amp;')
      expect(result).toContain('&lt;')
    })

    it('returns empty paragraph for empty input', () => {
      expect(markdownToNotesXml('')).toContain('<a:endParaRPr')
      expect(markdownToNotesXml('  ')).toContain('<a:endParaRPr')
    })

    it('joins continuation lines into one paragraph', () => {
      const result = markdownToNotesXml('Line one\nLine two')
      const count = (result.match(/<a:p>/g) || []).length
      expect(count).toBe(1)
      expect(result).toContain('Line one Line two')
    })
  })

  describe('extractNotesText', () => {
    it('extracts text from a notes slide XML', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>`
      expect(extractNotesText(xml)).toBe('Hello world')
    })

    it('preserves paragraph breaks as newlines', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>First paragraph</a:t></a:r></a:p>
      <a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>`
      expect(extractNotesText(xml)).toBe('First paragraph\nSecond paragraph')
    })

    it('returns null for notes without body placeholder', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
  </p:spTree></p:cSld>
</p:notes>`
      expect(extractNotesText(xml)).toBeNull()
    })

    it('returns null for empty notes text', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a:endParaRPr/></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>`
      expect(extractNotesText(xml)).toBeNull()
    })
  })

  describe('buildNotesSlideXml', () => {
    it('produces valid notes XML with three placeholders', () => {
      const xml = buildNotesSlideXml('<a:p><a:r><a:t>Test</a:t></a:r></a:p>')
      expect(xml).toContain('type="sldImg"')
      expect(xml).toContain('type="body"')
      expect(xml).toContain('type="sldNum"')
      expect(xml).toContain('<a:t>Test</a:t>')
    })
  })

  describe('buildNotesSlideRels', () => {
    it('produces rels pointing to the slide', () => {
      const xml = buildNotesSlideRels('slide1.xml')
      expect(xml).toContain('Target="../slides/slide1.xml"')
      expect(xml).toContain('relationships/slide')
    })
  })

  describe('buildNotesInjection', () => {
    it('creates notes files and updates rels', () => {
      const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`

      const files = buildNotesInjection(slideRels, 'Hello notes')
      expect(files['ppt/notesSlides/notesSlide1.xml']).toContain('<a:t>Hello notes</a:t>')
      expect(files['ppt/notesSlides/_rels/notesSlide1.xml.rels']).toContain('slide1.xml')
      expect(files['ppt/slides/_rels/slide1.xml.rels']).toContain('notesSlide')
      expect(files['ppt/slides/_rels/slide1.xml.rels']).toContain('rId2')
    })

    it('does not duplicate rels if notesSlide relationship already exists', () => {
      const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`

      const files = buildNotesInjection(slideRels, 'Updated notes')
      // Should not include slide rels since relationship already exists
      expect(files['ppt/slides/_rels/slide1.xml.rels']).toBeUndefined()
    })

    it('handles empty text to clear notes', () => {
      const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`

      const files = buildNotesInjection(slideRels, '')
      expect(files['ppt/notesSlides/notesSlide1.xml']).toContain('<a:endParaRPr')
    })
  })
})
