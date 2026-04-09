# OOXML Live Editing Reference

**Prerequisite**: Load the `/pptx` skill first for OOXML structure knowledge (namespaces, `<a:p>/<a:r>/<a:rPr>` anatomy, formatting rules, XML escaping). This reference covers how to apply that knowledge through the bridge's live-editing MCP tools.

## When to Use OOXML vs Office.js

| Approach | Best for |
|----------|----------|
| `format_shapes` | Fill color, font bold/italic/size/color/name — declarative batch formatting, 1 call per slide |
| Office.js (`execute_officejs`) | Shape creation, positioning, simple text writes — anything the Office.js API exposes directly |
| OOXML tools (`edit_slide_xml` with `code`) | Geometry (corners, borders), gradients, mixed formatting, precise paragraph control |
| OOXML tools (`edit_slide_xml` with `xml`) | Full slide XML replacement (rare — prefer code mode to avoid attribute loss) |
| `edit_shape_paragraphs` | Single-shape paragraph editing with OOXML (preserves bodyPr/lstStyle) |
| File-based (`get_local_copy` + `/pptx` skill) | Charts, master/theme editing, rels, Content_Types — anything beyond slide XML |

## Workflow

### Preferred: Code mode (1 call)

1. `inspect_slide(slideRange: "N")` → find shape IDs
2. `edit_slide_xml(slideIndex, code: "...")` → DOM manipulation server-side (read + modify in one call)
3. `screenshot_slide(slideIndex)` → visual verification

Code mode receives a pre-parsed DOM with helpers. See SKILL.md "OOXML Editing Workflow" for the full sandbox context and examples.

### Legacy: XML string mode (2+ calls)

1. `inspect_slide(slideRange: "N")` → find shape IDs
2. `read_shape_paragraphs(slideIndex, shapeId)` or `read_slide_xml(slideIndex, shapeId?)`
3. Modify the XML using `/pptx` skill knowledge
4. `edit_shape_paragraphs(slideIndex, shapeId, xml)` or `edit_slide_xml(slideIndex, xml, shapeId?)`
5. `screenshot_slide(slideIndex)` → visual verification

Always inspect before modifying. Always verify after.

## Shape ID Mapping

- `inspect_slide` returns shapes with an `id` field (e.g. `"5"`)
- This matches `<p:cNvPr id="5">` in the OOXML
- Always use `inspect_slide` first to discover IDs — don't guess
- Shape IDs may change after reimport (Office.js assigns new IDs on `insertSlidesFromBase64`)

## read_shape_paragraphs / edit_shape_paragraphs

Paragraph-level editing for a single shape:

- `read_shape_paragraphs` returns the `<a:p>` paragraph elements from a shape
- `edit_shape_paragraphs` replaces paragraph content — `<a:bodyPr>` and `<a:lstStyle>` are preserved automatically
- You only work with the paragraph XML (the `<a:p>` elements)

```
// Read current paragraphs
read_shape_paragraphs(slideIndex: 0, shapeId: "2")
// Returns: <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r></a:p>

// Write modified paragraphs back
edit_shape_paragraphs(slideIndex: 0, shapeId: "2", xml: '<a:p>..modified..</a:p>')
```

## read_slide_xml / edit_slide_xml

Full slide or shape-level XML editing:

- **Without shapeId**: returns/replaces the full slide XML (`<p:sld>...</p:sld>`)
- **With shapeId**: returns/replaces that shape's `<p:sp>` element only

```
// Full slide XML
read_slide_xml(slideIndex: 0)
edit_slide_xml(slideIndex: 0, xml: '<p:sld>...</p:sld>')

// Single shape XML
read_slide_xml(slideIndex: 0, shapeId: "5")
edit_slide_xml(slideIndex: 0, shapeId: "5", xml: '<p:sp>...</p:sp>')
```

Use full-slide mode for batch editing multiple shapes in a single reimport.

## Batching Multiple Edits

Each edit tool call triggers a full export → modify → delete → reimport cycle:

- Multiple `edit_shape_paragraphs` calls on the same slide = multiple reimports (visible flashing)
- **For 2+ shapes on the same slide**: use `read_slide_xml` (full slide, no shapeId) → modify all shapes in the XML → `edit_slide_xml` (full slide) — single reimport, no flashing

```
// Bad: 3 reimports, visible flashing
edit_shape_paragraphs(slideIndex: 0, shapeId: "2", xml: '...')
edit_shape_paragraphs(slideIndex: 0, shapeId: "5", xml: '...')
edit_shape_paragraphs(slideIndex: 0, shapeId: "7", xml: '...')

// Good: 1 reimport, no flashing
xml = read_slide_xml(slideIndex: 0)          // full slide
// modify shapes 2, 5, 7 in the XML
edit_slide_xml(slideIndex: 0, xml: modified)  // single reimport
```

## Units: Points vs EMU

| Context | Unit | 1 inch = |
|---------|------|----------|
| `inspect_slide` / Office.js | Points | 72 pt |
| OOXML (`<a:off>`, `<a:ext>`) | EMU | 914,400 EMU |

Conversion: **EMU = points × 12,700**

| Reference | Points | EMU |
|-----------|--------|-----|
| Standard 16:9 slide width | 960 pt | 12,192,000 |
| Standard 16:9 slide height | 540 pt | 6,858,000 |
| 1 inch | 72 pt | 914,400 |
| 1 cm | 28.35 pt | 360,000 |

When moving positions from `inspect_slide` output into OOXML, multiply by 12,700.

## Export/Reimport Mechanics

The bridge handles the export/reimport cycle transparently — you just send/receive XML. Under the hood:

1. **Export**: `slide.exportAsBase64()` → single-slide .pptx as Base64
2. **Unzip**: Server extracts `ppt/slides/slide1.xml` (always this path, regardless of slideIndex)
3. **Modify**: Server applies your XML changes to the extracted slide
4. **Repack**: Server creates a new Base64 .pptx with the modified XML
5. **Delete**: Original slide is deleted from the presentation
6. **Reimport**: `presentation.insertSlidesFromBase64()` at the same position

The data stays server-side — XML content never enters Claude's context.

## Multi-File Zip Access: read_slide_zip / edit_slide_zip

`read_slide_xml` / `edit_slide_xml` only access slide XML. For charts, rels, or Content_Types, use the zip-level tools:

```
// Discover all files in the exported zip
read_slide_zip(slideIndex: 0)
// Returns: { zipContents: { path: content, ... }, allPaths: [...] }

// Read specific files
read_slide_zip(slideIndex: 0, paths: ["ppt/charts/chart1.xml", "ppt/slides/_rels/slide1.xml.rels"])

// Update multiple files in one reimport (can add new files)
edit_slide_zip(slideIndex: 0, files: {
  "ppt/slides/slide1.xml": modifiedSlideXml,
  "ppt/charts/chart1.xml": chartXml,
  "ppt/slides/_rels/slide1.xml.rels": updatedRels
})
```

**Auto Content_Types**: When `edit_slide_zip` adds new files under `ppt/charts/`, it auto-registers them in `[Content_Types].xml`. You can still include `[Content_Types].xml` explicitly in the files map to override.

## Current Tool Limitations

The zip-level tools access all text/XML files in the **single-slide export**. They **cannot** access:

- **Masters/themes** — not included in single-slide export (need full pptx)
- **Binary media** — `ppt/media/` files are binary, not text (use `insert_image` instead)
- **Notes** — `ppt/notesSlides/` may not be included in single-slide export

**Workaround**: Use `get_local_copy` to get a .pptx file path, then edit with the `/pptx` skill's file-based workflow.

## Charts via OOXML

Use `read_slide_zip` / `edit_slide_zip` for chart creation and editing. Chart creation requires:

### Chart XML structure (`ppt/charts/chartN.xml`)

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <!-- Chart type element: barChart, lineChart, pieChart, etc. -->
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>
```

### Chart types

| OOXML element | Chart type |
|---------------|------------|
| `<c:barChart>` | Bar/column |
| `<c:lineChart>` | Line |
| `<c:pieChart>` | Pie |
| `<c:areaChart>` | Area |
| `<c:scatterChart>` | Scatter/XY |
| `<c:doughnutChart>` | Doughnut |

### Registration required

1. **Content_Types** — `edit_slide_zip` auto-registers this when adding `ppt/charts/*.xml` files
2. **Slide rels** — add `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>` to `ppt/slides/_rels/slide1.xml.rels`
3. **Graphic frame** on slide — add `<p:graphicFrame>` with `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId3"/></a:graphicData></a:graphic>` to the slide XML

### Chart Rules

- Every chart MUST include: `<c:title>`, `<c:legend>` (with `legendPos val="t"` and `overlay val="0"`), `<c:dLbls>` on every series
- Always include `<c:style val="2"/>` in `<c:chartSpace>` for theme color inheritance
- Do NOT hardcode series colors with `<c:spPr>` — omit so theme accents apply
- Stacked bar/column: MUST add `<c:overlap val="100"/>` inside `<c:barChart>` — without it, bars render side by side
- Category axis: `<c:majorTickMark val="none"/>` (ticks fall between categories and look offset)
- Value axis: `<c:majorTickMark val="out"/>` (keep major ticks visible)
- Both axes: `<c:minorTickMark val="none"/>`
- Chart font minimums: title sz="1600" (16pt), axis labels sz="1400" (14pt), data labels sz="1400" (14pt), legend sz="1400" (14pt)
- Pie/doughnut: add `<c:showPercent val="1"/>` and optionally `<c:showCatName val="1"/>` to `<c:dLbls>`

### Data Labels

Always add to each `<c:ser>`:
```xml
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>
```

### Full Chart Creation Workflow (8 Steps)

Complete executable workflow using `edit_slide_zip`. Each step builds on the previous.

```javascript
// Step 1: Read slide XML (always slide1.xml in exported zip)
var slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
var slideDoc = new DOMParser().parseFromString(slideXml, "text/xml");
var spTree = slideDoc.getElementsByTagName("p:spTree")[0];

// Step 2: Find next chart number
var chartFiles = Object.keys(zip.files).filter(function(f) { return /^ppt\/charts\/chart\d+\.xml$/.test(f); });
var chartNum = chartFiles.length + 1;

// Step 3: Build chart XML (no HTML/XML comments inside the string)
var chartXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
  + '              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + '              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  + '<c:style val="2"/>'
  + '<c:chart>'
  + '  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
  + '    <a:p><a:r><a:rPr lang="en-US" sz="1600" dirty="0"/><a:t>Chart Title</a:t></a:r></a:p>'
  + '  </c:rich></c:tx><c:overlay val="0"/></c:title>'
  + '  <c:plotArea>'
  + '    <c:barChart>'
  + '      <c:barDir val="col"/><c:grouping val="clustered"/>'
  + '      <c:ser>'
  + '        <c:idx val="0"/><c:order val="0"/>'
  + '        <c:cat><c:strLit><c:ptCount val="3"/>'
  + '          <c:pt idx="0"><c:v>Cat A</c:v></c:pt>'
  + '          <c:pt idx="1"><c:v>Cat B</c:v></c:pt>'
  + '          <c:pt idx="2"><c:v>Cat C</c:v></c:pt>'
  + '        </c:strLit></c:cat>'
  + '        <c:val><c:numLit><c:ptCount val="3"/>'
  + '          <c:pt idx="0"><c:v>42</c:v></c:pt>'
  + '          <c:pt idx="1"><c:v>78</c:v></c:pt>'
  + '          <c:pt idx="2"><c:v>55</c:v></c:pt>'
  + '        </c:numLit></c:val>'
  + '        <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>'
  + '      </c:ser>'
  + '      <c:axId val="1"/><c:axId val="2"/>'
  + '    </c:barChart>'
  + '    <c:catAx>'
  + '      <c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/>'
  + '      <c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
  + '      <c:tickLblPos val="nextTo"/><c:crossAx val="2"/>'
  + '      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" lang="en-US"/></a:pPr></a:p></c:txPr>'
  + '    </c:catAx>'
  + '    <c:valAx>'
  + '      <c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/>'
  + '      <c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
  + '      <c:tickLblPos val="nextTo"/><c:crossAx val="1"/>'
  + '      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" lang="en-US"/></a:pPr></a:p></c:txPr>'
  + '    </c:valAx>'
  + '  </c:plotArea>'
  + '  <c:legend>'
  + '    <c:legendPos val="t"/><c:overlay val="0"/>'
  + '    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" lang="en-US"/></a:pPr></a:p></c:txPr>'
  + '  </c:legend>'
  + '</c:chart></c:chartSpace>';

// Step 4: Add chart file to zip
zip.file("ppt/charts/chart" + chartNum + ".xml", chartXml);

// Step 5: Register in [Content_Types].xml
var ctXml = await zip.file("[Content_Types].xml").async("string");
var ctDoc = new DOMParser().parseFromString(ctXml, "text/xml");
var override = ctDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override");
override.setAttribute("PartName", "/ppt/charts/chart" + chartNum + ".xml");
override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
ctDoc.documentElement.appendChild(override);
zip.file("[Content_Types].xml", new XMLSerializer().serializeToString(ctDoc));

// Step 6: Add relationship in slide rels (always slide1.xml.rels)
var relsPath = "ppt/slides/_rels/slide1.xml.rels";
var relsXml = await zip.file(relsPath).async("string");
if (!relsXml) relsXml = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
var relsDoc = new DOMParser().parseFromString(relsXml, "text/xml");
var rId = "rId" + (relsDoc.getElementsByTagName("Relationship").length + 1);
var rel = relsDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
rel.setAttribute("Id", rId);
rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
rel.setAttribute("Target", "../charts/chart" + chartNum + ".xml");
relsDoc.documentElement.appendChild(rel);
zip.file(relsPath, new XMLSerializer().serializeToString(relsDoc));

// Step 7: Add graphic frame to slide — use DOMParser, NOT innerHTML (namespaces break with innerHTML)
var frameXml = '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + '<p:nvGraphicFramePr><p:cNvPr id="' + (chartNum + 100) + '" name="Chart ' + chartNum + '"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
  + '<p:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="7000000" cy="4000000"/></p:xfrm>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + '<c:chart r:id="' + rId + '"/>'
  + '</a:graphicData></a:graphic></p:graphicFrame>';
var frameDoc = new DOMParser().parseFromString(frameXml, "text/xml");
spTree.appendChild(slideDoc.importNode(frameDoc.documentElement, true));

// Step 8: Write slide back
zip.file("ppt/slides/slide1.xml", new XMLSerializer().serializeToString(slideDoc));
markDirty();
```

**Multi-series charts:** Repeat `<c:ser>` with incrementing `c:idx` and `c:order` values. Each series gets its own `<c:cat>` and `<c:val>`.

**Using `edit_slide_zip` (MCP tool approach):** The workflow above runs inside `edit_slide_zip`'s code environment. Alternatively, prepare all files externally and pass them to `edit_slide_zip`:

```
edit_slide_zip(slideIndex: 0, files: {
  "ppt/slides/slide1.xml": modifiedSlideXml,
  "ppt/charts/chart1.xml": chartXml,
  "ppt/slides/_rels/slide1.xml.rels": updatedRelsXml
})
```

`edit_slide_zip` auto-registers Content_Types for new `ppt/charts/*.xml` files.

## Master/Theme via OOXML (Future Reference)

When full pptx export is added, theme editing involves:

### Theme structure (`ppt/theme/theme1.xml`)

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <!-- accent2-6, hlink, folHlink -->
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
```

### What execute_officejs CAN do now

- Set slide backgrounds: `slide.fill.setSolidColor("#hex")`
- Read some theme colors via shapes
- Apply formatting that references theme colors

### What needs full pptx export

- Full color scheme editing (`<a:clrScheme>`)
- Font scheme changes (`<a:fontScheme>`)
- Master slide layout editing (`ppt/slideMasters/`, `ppt/slideLayouts/`)
- Background styles with gradients or patterns

## Text Formatting Rules

### The Core Rule

**Never use Office.js to read text content.** `textRange.text` returns plain text — all formatting (bold, font size, color, bullets) is stripped. Use `read_shape_paragraphs` to read. Office.js is only for shape metadata (IDs, positions, dimensions) and simple plain-text writes.

### OOXML Text Elements Reference

| Element | Purpose | Example |
|---------|---------|---------|
| `<a:p>` | Paragraph | Container for runs |
| `<a:r>` | Text run | Container for text + properties |
| `<a:t>` | Text content | `<a:t>Hello</a:t>` |
| `<a:rPr>` | Run properties | `b="1"` bold, `i="1"` italic, `sz="1600"` = 16pt, `u="sng"` underline. Always include `lang="en-US" dirty="0"`. |
| `<a:pPr>` | Paragraph properties | `algn="ctr\|l\|r\|just"`, `lvl="1"` indent level (0-based) |
| `<a:latin>` | Font family (inside `<a:rPr>`) | `typeface="Calibri"` |
| `<a:solidFill>` | Color (inside `<a:rPr>`) | `<a:srgbClr val="FF0000"/>` (hex) or `<a:schemeClr val="accent1"/>` (theme) |
| `<a:lnSpc>` | Line spacing (inside `<a:pPr>`) | `<a:spcPts val="1400"/>` (hundredths of pt) |
| `<a:spcBef>` / `<a:spcAft>` | Space before/after paragraph | `<a:spcPts val="600"/>` |

### Bullet and List Patterns

**Top-level bullet:**
```xml
<a:p>
  <a:pPr><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>
  <a:r><a:rPr lang="en-US" sz="1600" dirty="0"/><a:t>Bullet text</a:t></a:r>
</a:p>
```

**Sub-bullet** (`lvl="1"` = second level, 0-based):
```xml
<a:p>
  <a:pPr lvl="1"><a:buFont typeface="Arial"/><a:buChar char="&#8211;"/></a:pPr>
  <a:r><a:rPr lang="en-US" sz="1400" dirty="0"/><a:t>Sub-bullet</a:t></a:r>
</a:p>
```

**Numbered list:**
```xml
<a:p>
  <a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>
  <a:r><a:rPr lang="en-US" sz="1600" dirty="0"/><a:t>Numbered item</a:t></a:r>
</a:p>
```

**Header (no bullet):**
```xml
<a:p>
  <a:pPr><a:buNone/></a:pPr>
  <a:r><a:rPr lang="en-US" sz="1800" b="1" dirty="0"/><a:t>Header text</a:t></a:r>
</a:p>
```

### Bullet Rules

- `lvl` is 0-based. Top-level bullets = `lvl="0"` or omit `lvl` (default is 0). Sub-bullets = `lvl="1"`.
- Headers are NOT a "level" — they are level 0 with `<a:buNone/>`.
- NEVER put the bullet character in `<a:t>` — use `<a:buChar char="..."/>` in `<a:pPr>`.
- NEVER use `lvl="1"` for top-level bullets.
- When editing, copy the existing `<a:pPr>` (which may use explicit `marL`/`indent` instead of `lvl`) rather than inventing new attributes.
- New bullets must copy bullet `<a:pPr>`, not header `<a:pPr>`.

### Hanging Indent Pattern

Some templates use explicit `marL`/`indent` attributes instead of `lvl` for bullet indentation:

```xml
<a:pPr marL="228600" indent="-228600">
  <a:buFont typeface="Arial"/>
  <a:buChar char="&#8226;"/>
</a:pPr>
```

`marL="228600"` = left margin in EMUs (indents the whole paragraph including wrapped lines). `indent="-228600"` = first-line indent in EMUs (outdents the bullet character back to the left edge). Together they create a hanging indent. When editing, copy the existing `<a:pPr>` verbatim rather than switching to `lvl`.

### AutoSize Options

| Setting | Behavior | Use when |
|---------|----------|----------|
| `"AutoSizeShapeToFitText"` | Shape expands/contracts to fit text | Layout-critical text boxes; MUST set before `verify_slides` |
| `"AutoSizeTextToFitShape"` | Text shrinks to fit fixed shape size | Fixed-size containers |
| `"AutoSizeNone"` | Fixed size, no adjustment | Centered-text shapes, icon labels |

### DOs

- Always call `read_shape_paragraphs` before `edit_shape_paragraphs` to see the existing XML
- Copy every `<a:p>` block verbatim from the read output, then make only the specific change needed
- Copy formatting from similar paragraphs when adding new content — new bullets should use the same `<a:pPr>` and `<a:rPr>` as existing ones
- Use `<a:buChar>` in `<a:pPr>` for native PowerPoint bullets
- Keep theme colors (`<a:schemeClr>`) — never replace with hardcoded hex unless explicitly asked
- Escape XML special characters in `<a:t>`: `&` → `&amp;`, `<` → `&lt;`, `>` → `&gt;`

### DON'Ts

- **Don't put bare `&` in `<a:t>`** — always escape as `&amp;`. This is the #1 cause of missing text. `Sales & Marketing` must be `Sales &amp; Marketing`
- Don't rewrite or "clean up" XML — copy verbatim. If read returns `<a:rPr lang="en-US" sz="1000" b="1" dirty="0">`, write exactly that
- Don't use `lvl="1"` for top-level bullets — lvl is 0-based: top-level = `lvl="0"` or omit lvl. `lvl="1"` creates sub-bullets
- When editing, copy existing `<a:pPr>` (which may use explicit marL/indent) rather than inventing new attributes
- Don't put the `•` character in `<a:t>` — use `<a:buChar char="•"/>` in `<a:pPr>`
- Don't mix header and bullet formatting — headers use `<a:buNone/>` with different attributes

### edit_shape_paragraphs Preserves bodyPr and lstStyle

`edit_shape_paragraphs` automatically preserves the shape's `<a:bodyPr>` and `<a:lstStyle>` — you only provide the `<a:p>` paragraph elements. This means text anchoring, margins, and list style definitions carry over automatically. The `replaceTextBody` helper below mirrors this behavior for batch edits via `edit_slide_xml`.

### Batch Edit Helpers

When editing 2+ shapes on the same slide, use full-slide `read_slide_xml` / `edit_slide_xml` with these helpers to avoid multiple reimport flashes:

```javascript
var NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
var NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function findShapeById(doc, id) {
  var shapes = doc.getElementsByTagNameNS(NS_P, "sp");
  for (var i = 0; i < shapes.length; i++) {
    var nvSpPr = shapes[i].getElementsByTagNameNS(NS_P, "nvSpPr")[0];
    var cNvPr = nvSpPr ? nvSpPr.getElementsByTagNameNS(NS_P, "cNvPr")[0] : null;
    if (cNvPr && cNvPr.getAttribute("id") === id) return shapes[i];
  }
  return null;
}

function replaceTextBody(doc, shape, paragraphXml) {
  var txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
  if (!txBody) return;
  // Preserve bodyPr and lstStyle (same as edit_shape_paragraphs behavior)
  var bodyPr = txBody.getElementsByTagNameNS(NS_A, "bodyPr")[0];
  var lstStyle = txBody.getElementsByTagNameNS(NS_A, "lstStyle")[0];
  while (txBody.firstChild) txBody.removeChild(txBody.firstChild);
  if (bodyPr) txBody.appendChild(bodyPr);
  if (lstStyle) txBody.appendChild(lstStyle);
  var wrapper = new DOMParser().parseFromString(
    '<w xmlns:a="' + NS_A + '">' + paragraphXml + '</w>', "text/xml"
  ).documentElement;
  for (var i = 0; i < wrapper.childNodes.length; i++) {
    if (wrapper.childNodes[i].nodeType === 1)
      txBody.appendChild(doc.importNode(wrapper.childNodes[i], true));
  }
}

// Usage in edit_slide_xml:
var xml = await zip.file("ppt/slides/slide1.xml").async("string");
var doc = new DOMParser().parseFromString(xml, "text/xml");

var title = findShapeById(doc, "2");
replaceTextBody(doc, title, '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1" dirty="0"/><a:t>New Title</a:t></a:r></a:p>');

var body = findShapeById(doc, "3");
replaceTextBody(doc, body, '<a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr><a:r><a:rPr lang="en-US" sz="1600" dirty="0"/><a:t>Point 1</a:t></a:r></a:p>');

zip.file("ppt/slides/slide1.xml", new XMLSerializer().serializeToString(doc));
markDirty();
```

Always find shapes by ID (not name — names are locale-dependent). Shape IDs come from `inspect_slide` output. The exported slide is always `ppt/slides/slide1.xml` in the zip regardless of `slideIndex`.

## Diagrams and Infographics via OOXML

Use `edit_slide_xml` for process flows, timelines, cycles, org charts, and custom layouts. OOXML gives precise control over shape positioning, text anchoring, and alignment that Office.js cannot match for these layouts.

### OOXML Units (EMUs)

1 inch = 914,400 EMU. 1 point = 12,700 EMU. Standard slide: 9,144,000 x 6,858,000 EMU (960 x 540 pt).

| Points | EMU |
|--------|-----|
| 50 pt | 635,000 |
| 100 pt | 1,270,000 |
| 480 pt (center X) | 6,096,000 |
| 960 pt (full width) | 12,192,000 |

### Shape Element Structure

```xml
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="10" name="Box1"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="1500000" cy="500000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="2F5496"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r>
      <a:rPr lang="en-US" sz="1400" b="1" dirty="0">
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </a:rPr>
      <a:t>Step 1</a:t>
    </a:r></a:p>
  </p:txBody>
</p:sp>
```

Key elements: `<p:nvSpPr>` (identity), `<p:spPr>` (position + geometry + fill), `<p:txBody>` (text content). Position via `<a:off x="" y=""/>`, size via `<a:ext cx="" cy=""/>` (all in EMU). Preset geometry via `<a:prstGeom prst="roundRect">`.

### Connector Element Structure

```xml
<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvCxnSpPr>
    <p:cNvPr id="20" name="Connector1"/>
    <p:cNvCxnSpPr/><p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="2500000" y="1200000"/><a:ext cx="800000" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>
```

For arrows, use `<a:tailEnd type="triangle"/>` inside `<a:ln>`. Line width `w` is in EMU (25400 = 2pt).

### escapeXml for User Text

Always use `escapeXml(text)` (available as a global in `edit_slide_xml` and `edit_slide_chart`) when embedding user text in XML template strings. A bare `&` or `<` in `<a:t>` breaks the XML parser and silently drops all subsequent text runs.

```javascript
var label = escapeXml("R&D / Innovation");
// Result: "R&amp;D / Innovation"
// Usage: '...<a:t>' + escapeXml(userText) + '</a:t>...'
```

Escapes: `&` -> `&amp;`, `<` -> `&lt;`, `>` -> `&gt;`, `"` -> `&quot;`, `'` -> `&apos;`

## Hyperlinks in OOXML

Adding a hyperlink requires BOTH an XML attribute in the run properties AND a relationship entry in the rels file. Use `edit_slide_zip` since you need to modify two files.

**Step 1:** Add `<a:hlinkClick>` to the run properties in the slide XML:

```xml
<a:r>
  <a:rPr lang="en-US" sz="1600" dirty="0">
    <a:hlinkClick r:id="rId5"/>
  </a:rPr>
  <a:t>Click here</a:t>
</a:r>
```

**Step 2:** Register the relationship in `ppt/slides/_rels/slide1.xml.rels`:

```xml
<Relationship
  Id="rId5"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com"
  TargetMode="External"/>
```

`TargetMode="External"` is required for URLs. The `rId` in `<a:hlinkClick>` must match the `Id` in the `<Relationship>` element. Follow the same rels-file pattern as adding chart relationships — read, parse with DOMParser, add Relationship element, serialize, write back.

## Pipeline-Specific Gotchas

1. **Shape IDs change after reimport** — Office.js assigns new IDs on `insertSlidesFromBase64`. Always re-read `inspect_slide` after editing if you need to reference shapes again.

2. **Edit slides in reverse index order** — each reimport deletes and reinserts the slide. If editing slides 0, 1, 2, edit in order 2 → 1 → 0 to avoid index shifting.

3. **Namespace variations** — `read_slide_xml` returns the exported slide's XML verbatim. Namespace prefixes may differ slightly from a raw .pptx file due to Office.js export behavior. Match what you read, don't assume canonical prefixes.

4. **Single-slide export scope** — the exported zip always contains just one slide at `ppt/slides/slide1.xml`, even if the original was slide 5 in the deck. Shape references to external content (hyperlinks, charts, media) may break if rels aren't included.

5. **Reimport is destructive** — the original slide is deleted before reimport. If the modified XML is malformed, the slide may be lost. Always keep the read XML as a fallback reference.

6. **Concurrent edits** — if the user edits the slide in PowerPoint while you're modifying XML, the reimport will overwrite their changes. Warn users before batch OOXML operations.
