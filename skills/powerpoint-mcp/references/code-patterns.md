# Office.js Code Patterns

All code passed to `execute_officejs` runs inside `PowerPoint.run(async (context) => { ... })`.
The `context` variable is pre-bound. Always call `await context.sync()` after loading properties or batching changes.

## Slides

```javascript
// Add blank slide
context.presentation.slides.add();
await context.sync();

// Delete slide by index
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
slides.items[2].delete();
await context.sync();

// Get slide count
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
return slides.items.length;
```

## Adding Slides with Layouts

**Preferred**: Use the `add_slide` tool — it handles layout lookup, positioning, and returns placeholder IDs in one call:
- `add_slide({ layoutName: "Title and Content", position: 2 })` → adds slide at index 2 with that layout, returns `{ slideIndex, placeholders: [{ id, name, type }] }`

**Manual approach** (only if `add_slide` is unavailable): use `execute_officejs` with the code below.

Always pick the layout that best matches content. Do NOT use "Blank" for slides with text.

Common layouts: "Title Slide", "Title and Content", "Two Content", "Section Header", "Title Only", "Blank"

### Layout Selection Rules

- NEVER use Blank layout for text-heavy slides -- no placeholders means no inherited font sizes, leading to inconsistent rendering.
- Blank is ONLY for fully custom visual slides (full-bleed images, pure-shape diagrams) with no text structure.
- Layout purpose cheat sheet:
  - "Title Slide" -- opening slide / section dividers (large centered title + subtitle)
  - "Title and Content" -- standard body slide with title + content placeholder
  - "Two Content" -- side-by-side comparison
  - "Comparison" -- two columns with headers for comparing options
  - "Content with Caption" -- main content area + smaller sidebar/description area
  - "Section Header" -- transition slides between major sections
  - "Title Only" -- title + custom shapes below

```javascript
// Find the right layout (always use the last master — earlier may be stale)
var masters = context.presentation.slideMasters;
masters.load("items");
await context.sync();
var slideMaster = masters.items[masters.items.length - 1];
slideMaster.layouts.load("items,name");
await context.sync();
var layout = slideMaster.layouts.items.find(function(l) { return l.name === "Title and Content"; });

// Add a slide with that layout
var slides = context.presentation.slides;
slides.add({ layoutId: layout.id });
await context.sync();
slides.load("items");
await context.sync();
var newSlide = slides.items[slides.items.length - 1];

// Use placeholders (find by name pattern, then use shape.id for subsequent calls)
var shapes = newSlide.shapes;
shapes.load("items/id,items/name");
await context.sync();
var entries = shapes.items.map(function(shape) {
  var tf = shape.getTextFrameOrNullObject();
  tf.load(["hasText", "textRange"]);
  return { shape: shape, tf: tf };
});
await context.sync();

var title = entries.find(function(e) { return !e.tf.isNullObject && e.shape.name.startsWith("Title"); });
if (title) title.tf.textRange.text = "My Slide Title";

// Reposition: slides.add() always appends
newSlide.moveTo(4); // 0-based target index
await context.sync();
```

**Don't delete "empty" placeholders after writing text.** The `hasText` you loaded is stale once you set `textRange.text`, so a cleanup loop will delete what you just wrote.

To change layout on existing slide: `slide.applyLayout(layout);`

## Building Entire Slides in One Call

For efficiency, create ALL shapes for a slide in a single `execute_officejs` call. This avoids multiple round-trips, prevents mid-build visual flashing, and is much faster than separate calls per shape.

```javascript
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
var shapes = slides.items[0].shapes;

// Card 1: background + title + body
var card1 = shapes.addGeometricShape("RoundRectangle", { left: 50, top: 200, width: 420, height: 280 });
card1.fill.setSolidColor("#F5F0EB");
var title1 = shapes.addTextBox("Feature One", { left: 135, top: 220, width: 320, height: 30 });
title1.textFrame.textRange.font.size = 18;
title1.textFrame.textRange.font.bold = true;
var body1 = shapes.addTextBox("Description of the first feature with key benefits.", { left: 135, top: 260, width: 320, height: 100 });
body1.textFrame.textRange.font.size = 14;
body1.textFrame.wordWrap = true;

// Card 2: same pattern, offset right
var card2 = shapes.addGeometricShape("RoundRectangle", { left: 500, top: 200, width: 420, height: 280 });
card2.fill.setSolidColor("#F5F0EB");
var title2 = shapes.addTextBox("Feature Two", { left: 585, top: 220, width: 320, height: 30 });
title2.textFrame.textRange.font.size = 18;
title2.textFrame.textRange.font.bold = true;
var body2 = shapes.addTextBox("Description of the second feature.", { left: 585, top: 260, width: 320, height: 100 });
body2.textFrame.textRange.font.size = 14;
body2.textFrame.wordWrap = true;

// Divider line (use slideWidth from inspect_deck/inspect_slide response)
shapes.addLine("Straight", { left: 0, top: 180, width: slideWidth, height: 0 });

await context.sync();
```

Build all shapes in one batch, then call `context.sync()` once at the end.

## Template Chrome Awareness

Template slides typically have placeholder shapes from the slide layout (metadata bar, title, slide number, footer) and sometimes decorative elements (HR divider lines). Before building content:

1. List existing shapes via `inspect_slide` or `scan_slide` to see what placeholders exist
2. Note the placeholder positions — especially the title area (typically top ~108pt) and footer area (bottom ~763pt)
3. Build your content shapes BELOW the existing chrome elements, not overlapping them
4. Set text in existing placeholders (title, subtitle) rather than creating new TextBoxes for those roles

```javascript
// Discover existing placeholders before adding content
var shapes = slides.items[0].shapes;
shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
await context.sync();
var placeholders = shapes.items.filter(function(s) { return s.type === "Placeholder"; });
// placeholders tell you where the chrome is — build content below them
```

## Geometric Shapes

```javascript
var slides = context.presentation.slides;
slides.load("items");
await context.sync();
var shapes = slides.items[0].shapes;

// Add rectangle
var rect = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
rect.left = 100; rect.top = 100; rect.width = 200; rect.height = 150;
rect.fill.setSolidColor("#2196F3");
await context.sync();
```

Use string literals for shape types:

```javascript
var shape = shapes.addGeometricShape("Rectangle", {
  left: 100, top: 100, width: 200, height: 100
});
```

Valid shape types (exact strings):
- **Basic:** Rectangle, RoundRectangle, Triangle, RightTriangle, Diamond, Parallelogram, Trapezoid, Pentagon, Hexagon, Octagon
- **Curved:** Ellipse, Donut, Arc, Pie, Chevron, HomePlate, Teardrop, BlockArc
- **Arrows:** RightArrow, LeftArrow, UpArrow, DownArrow, LeftRightArrow, UpDownArrow, BentArrow, CurvedRightArrow, CurvedLeftArrow, CircularArrow, StripedRightArrow, NotchedRightArrow
- **Callouts:** WedgeRectCallout, WedgeRRectCallout, WedgeEllipseCallout, CloudCallout, Cloud
- **Flowchart:** FlowChartProcess, FlowChartDecision, FlowChartInputOutput, FlowChartDocument, FlowChartTerminator, FlowChartConnector, FlowChartAlternateProcess
- **Brackets:** BracketPair, BracePair, LeftBracket, RightBracket, LeftBrace, RightBrace
- **Stars:** Star4, Star5, Star6, Star8
- **Other:** Plus, Frame, Funnel, Cube, Heart, LightningBolt

## Text Boxes & Lines

```javascript
// Text box with position
shapes.addTextBox("Hello World", { left: 50, top: 50, width: 300, height: 40 });
await context.sync();

// Straight line
shapes.addLine(PowerPoint.ConnectorType.straight, {
  left: 50, top: 200, width: 400, height: 0
});
await context.sync();
```

## Text Formatting

```javascript
// Set text and color
shape.textFrame.textRange.text = "Styled text";
shape.textFrame.textRange.font.color = "#FF5722";
shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
await context.sync();

// Multiline
shape.textFrame.textRange.text = "Line 1\nLine 2\nLine 3";
await context.sync();
```

## Safe Text Frame Access

Not all shapes have a text frame — tables, images, charts, grouped shapes throw `InvalidArgument` if you access `.textFrame` directly. **Always use `getTextFrameOrNullObject()`:**

```javascript
shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
await context.sync();

var textFrames = shapes.items.map(function(shape) {
  var tf = shape.getTextFrameOrNullObject();
  tf.load(["hasText"]);
  return { shape: shape, tf: tf };
});
await context.sync();

var textShapes = textFrames.filter(function(e) { return !e.tf.isNullObject; });
```

Never use `.textFrame` directly — use `getTextFrameOrNullObject()` and check `.isNullObject`.

### Stale Value Pitfall

All loaded properties are **snapshots** from the last `context.sync()`. They do NOT update after a write.

```javascript
// WRONG — hasText stays false even after writing
var tf = shape.getTextFrameOrNullObject();
tf.load("hasText");
await context.sync();
tf.getRange().text = "Hello";
await context.sync();
// tf.hasText is STILL false (stale snapshot from the first sync)
```

This applies to ALL loaded properties: `hasText`, `width`, `height`, `left`, `top`, `name`, `type`. To get fresh values after a write, re-load the properties and call `context.sync()` again. Do NOT branch on stale reads.

## Centering Text in Shapes

When placing text inside a geometric shape (numbers in circles, labels in rectangles), put text in the shape's own `textFrame` — never create a separate `addTextBox`. Set ALL of these:

```javascript
shape.textFrame.textRange.text = "$";
shape.textFrame.textRange.font.color = "#FFFFFF";
shape.textFrame.textRange.font.size = 16;
shape.textFrame.textRange.font.bold = true;
shape.textFrame.textRange.paragraphFormat.alignment = "Center";
shape.textFrame.verticalAlignment = "Middle";
shape.textFrame.autoSizeSetting = "AutoSizeNone";
shape.textFrame.wordWrap = false;
shape.textFrame.marginLeft = 0;
shape.textFrame.marginRight = 0;
shape.textFrame.marginTop = 0;
shape.textFrame.marginBottom = 0;
```

Missing any of these will cause off-center text. AutoSize options: `"AutoSizeShapeToFitText"` (shape expands to fit), `"AutoSizeTextToFitShape"` (text shrinks to fit), `"AutoSizeNone"` (fixed size).

### Text Box Inset Rule

Two distinct patterns for text placement -- do not confuse them:

1. **Text INSIDE its own shape** (centered label, icon number): use the shape's `textFrame` with `verticalAlignment = "Middle"`, `paragraphFormat.alignment = "Center"`, ALL margins = 0, `wordWrap = false`. See the centering pattern above.

2. **TextBox OVER a background shape** (card body text, banner label): inset the TextBox 10-15pt per side so text does not touch the background shape's edges:

```javascript
// Background shape at left=100, top=200, width=200, height=80
var bg = shapes.addGeometricShape("RoundRectangle", { left: 100, top: 200, width: 200, height: 80 });
// TextBox inset 12pt on each side
var tb = shapes.addTextBox("Card title", {
  left: 112,   // 100 + 12
  top: 212,    // 200 + 12
  width: 176,  // 200 - 24 (12 each side)
  height: 56   // 80 - 24
});
```

## Fill & Color

Only solid fills supported — no gradients, shadows, or effects.

```javascript
shape.fill.setSolidColor("#4CAF50");
// Named colors work too
shape.fill.setSolidColor("lightblue");
shape.fill.setSolidColor("coral");
await context.sync();
```

## Grouping

```javascript
var group = shapes.addGroup(arrayOfShapes);
await context.sync();

// Ungroup
group.group.ungroup();
await context.sync();
```

## Tables (API 1.8+)

```javascript
// Add table with values (every cell must be a string)
var shape = shapes.addTable(3, 4, {
  values: [
    ["Name", "Q1", "Q2", "Q3"],
    ["Alice", "100", "150", "120"],
    ["Bob", "90", "110", "140"]
  ],
  left: 100, top: 150, width: 500, height: 150
});
await context.sync();
```

### Formatting Cells

```javascript
var table = shape.getTable();
for (var col = 0; col < 4; col++) {
  var cell = table.getCellOrNullObject(0, col);
  cell.fill.setSolidColor("#2F5496");
  cell.font.color = "#FFFFFF";
  cell.font.bold = true;
  cell.font.size = 14;
  cell.horizontalAlignment = "Center";
  cell.verticalAlignment = "Middle";
}
await context.sync();
```

### Cell Properties

- `cell.text` — get/set text
- `cell.fill.setSolidColor(color)` — background
- `cell.font.bold`, `.italic`, `.size`, `.color`, `.name` — font
- `cell.horizontalAlignment` — "Left", "Center", "Right", "Justify"
- `cell.verticalAlignment` — "Top", "Middle", "Bottom"

### Merging, Rows, Columns

- `table.mergeCells(rowIndex, colIndex, rowCount, colCount)`
- `table.rows.add(index, count)` / `table.columns.add(index, count)`
- `table.columns.getItemAt(i).width = 200` / `table.rows.getItemAt(i).height = 40`
- Built-in styles: `"ThemedStyle1Accent1"` through `"ThemedStyle2Accent6"`, `"NoStyleTableGrid"`
- **Style is creation-time only** — the `style` property must be passed in `addTable()` options at creation time. It cannot be applied after the table exists:

```javascript
// Correct — style at creation time
var shape = shapes.addTable(3, 4, {
  values: [["A", "B", "C", "D"], ["1", "2", "3", "4"], ["5", "6", "7", "8"]],
  left: 80, top: 120, width: 800, height: 200,
  style: "ThemedStyle1Accent2"
});

// Wrong — there is no table.style property after creation
// table.style = "ThemedStyle1Accent2"; // does not exist in the API
```

### Table Rules

- Font minimum 14pt for all cells including headers
- Row height: 28-32pt single-line, 48-56pt two-line. Set table height to match row estimates.
- If any cell needs 3+ sentences or exceeds ~40 words — truncate, footnote, or split across slides
- **Table height is auto-calculated** — setting `shape.height` or OOXML `<a:ext cy>` is overridden by PowerPoint. Fix overflow via `cell.font.size` + `row.height` through the table API, not XML or shape properties.

### Table Content Suitability

- **Cell density limit:** if any cell needs 3+ sentences or exceeds ~40 words, the content is too dense. Either truncate to one concise sentence, move detail to a text box / footnote below the table, or split across slides.
- **Row height planning:** 28-32pt for single-line cells, 48-56pt for two-line cells. Three+ line cells should be split.
- Set table height to match row estimates. Do NOT constrain height first and then shrink fonts to fit.
- If total table height would exceed ~400pt (leaving room for title + margins), split the table across multiple slides.

## Deck Overview

Use the `preview_deck` MCP tool to review an entire presentation efficiently. Returns thumbnails interleaved with text metadata in one call — far cheaper than sequential `inspect_slide` + `screenshot_slide` per slide.

```
preview_deck()                              // all slides, thumbnails at 480px
preview_deck(slideRange: "0-5")             // first 6 slides only
preview_deck(slideRange: "2,4,7")           // specific slides
preview_deck(includeImages: false)           // text-only (fastest)
preview_deck(imageWidth: 720)                // larger thumbnails
```

## Screenshots

Always use the `screenshot_slide` MCP tool for visual screenshots. Do NOT call `getImageAsBase64` through `execute_officejs` — the raw Base64 text overflows the token limit.

```javascript
// Only if you need the raw Base64 data (prefer screenshot_slide tool instead)
var slide = slides.items[0];
var result = slide.getImageAsBase64({ width: 720 });
await context.sync();
return { base64: result.value };

// With custom dimensions (height auto-calculated if omitted)
var result = slide.getImageAsBase64({ width: 1280, height: 720 });
await context.sync();
```

## Inserting Images

Use the `insert_image` MCP tool to insert images onto slides. Three source modes:

```
insert_image(
  source: "/path/to/image.png",
  sourceType: "file",         // reads from disk, data stays server-side
  slideIndex: 0,              // optional: navigate to this slide first (0-based)
  left: 100,                  // optional: position in points
  top: 100,
  width: 400,
  height: 300,
)

insert_image(
  source: "https://example.com/photo.jpg",
  sourceType: "url",          // fetches from URL, data stays server-side
)

insert_image(
  source: "iVBORw0KGgo...",
  sourceType: "base64",       // raw base64 data
)
```

For `file` and `url` modes, image data transfers server-side and never enters Claude's context. If position/size are omitted, Office.js uses defaults.

## Copying Slides Between Presentations

Use the `copy_slides` MCP tool to copy slides between two open presentations. The Base64 data transfers server-side (Add-in A → Bridge Server → Add-in B) and never enters Claude's context.

```
copy_slides(
  sourceSlideIndex: 2,
  sourcePresentationId: "deck-a.pptx",
  destinationPresentationId: "deck-b.pptx",
  formatting: "UseDestinationTheme",  // optional
  targetSlideId: "267#"               // optional: insert after this slide
)
```

**Formatting options:**
- `"KeepSourceFormatting"` (default) — inserted slides keep their original theme/colors
- `"UseDestinationTheme"` — inserted slides adopt the target presentation's theme

**Slide ID formats** for `targetSlideId`:
- `"267#"` — slide ID only
- `"#763315295"` — creation ID only
- `"267#763315295"` — both

Under the hood, `copy_slides` calls `slide.exportAsBase64()` on the source and `presentation.insertSlidesFromBase64()` on the destination. For direct use via `execute_officejs`:

```javascript
// Export a slide to Base64 .pptx (API 1.8+)
var slide = slides.items[0];
var result = slide.exportAsBase64();
await context.sync();
return result.value; // Base64 .pptx string

// Insert slides from Base64 .pptx
context.presentation.insertSlidesFromBase64(base64String, {
  formatting: "UseDestinationTheme",
  targetSlideId: "267#"
});
await context.sync();

// Selective import: insert only specific slides from a multi-slide .pptx
// Use sourceSlideIds to pick which slides to import — do NOT insert all then delete extras
context.presentation.insertSlidesFromBase64(base64String, {
  formatting: "UseDestinationTheme",
  sourceSlideIds: ["256", "258"]  // slide IDs from the source .pptx
});
await context.sync();
```

## OOXML Text Editing

Use the OOXML tools for fine-grained formatting control. Load the `/pptx` skill for OOXML structure knowledge. See [ooxml-reference.md](ooxml-reference.md) for the full live-editing reference (batching, units, gotchas).

```
// 1. Read current paragraphs (raw OOXML)
read_shape_paragraphs(slideIndex: 0, shapeId: "2")
// Returns: <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r></a:p>

// 2. Modify the XML (add color, change text, etc.)
// 3. Write back
edit_shape_paragraphs(slideIndex: 0, shapeId: "2", xml: '<a:p><a:r><a:rPr lang="en-US" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Red Bold</a:t></a:r></a:p>')
```

For full slide or shape XML editing:

```
// Read full slide XML
read_slide_xml(slideIndex: 0)

// Read specific shape XML
read_slide_xml(slideIndex: 0, shapeId: "5")

// Replace a shape's XML
edit_slide_xml(slideIndex: 0, shapeId: "5", xml: '<p:sp>...</p:sp>')

// Replace full slide XML (batch multiple shapes in one reimport)
edit_slide_xml(slideIndex: 0, xml: '<p:sld>...</p:sld>')
```

## Multi-File Zip Editing

For charts, rels, Content_Types, or other zip entries beyond slide XML. See [ooxml-reference.md](ooxml-reference.md) for details.

```
// 1. Read all text/xml files from the slide zip (auto-discovers)
read_slide_zip(slideIndex: 0)
// Returns: { zipContents: { "ppt/slides/slide1.xml": "...", "[Content_Types].xml": "...", ... }, allPaths: [...] }

// 2. Read specific files
read_slide_zip(slideIndex: 0, paths: ["ppt/charts/chart1.xml", "ppt/slides/_rels/slide1.xml.rels"])

// 3. Update multiple files in one reimport
edit_slide_zip(slideIndex: 0, files: {
  "ppt/slides/slide1.xml": "<p:sld>...modified...</p:sld>",
  "ppt/charts/chart1.xml": "<c:chartSpace>...</c:chartSpace>"
})
// Auto-registers Content_Types for new chart files
```

## Duplicating Slides

```
// Duplicate slide 2 right after itself
duplicate_slide(slideIndex: 2)

// Duplicate slide 0, insert after slide 4
duplicate_slide(slideIndex: 0, insertAfter: 4)
```

## Charts (Declarative)

Use `edit_slide_chart` for automatic chart creation from structured data — no OOXML knowledge needed:

```
// Column chart
edit_slide_chart(
  slideIndex: 0,
  chartType: "column",
  title: "Revenue by Quarter",
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [
    { name: "2024", values: [100, 150, 120, 180] },
    { name: "2025", values: [130, 170, 140, 200] }
  ]
)

// Pie chart
edit_slide_chart(
  slideIndex: 1,
  chartType: "pie",
  title: "Market Share",
  categories: ["Product A", "Product B", "Other"],
  series: [{ name: "Share", values: [45, 35, 20] }]
)

// Stacked bar with custom position
edit_slide_chart(
  slideIndex: 2,
  chartType: "bar",
  title: "Team Capacity",
  categories: ["Dev", "Design", "QA"],
  series: [
    { name: "Allocated", values: [80, 60, 40] },
    { name: "Available", values: [20, 40, 60] }
  ],
  position: { left: 100, top: 120, width: 500, height: 350 },
  options: { stacked: true, legendPosition: "b" }
)
```

Supported chart types: `column`, `bar`, `line`, `pie`, `doughnut`, `area`. For advanced chart customization beyond what `edit_slide_chart` supports, use `edit_slide_zip` with raw OOXML (see [ooxml-reference.md](ooxml-reference.md)).

## Verifying Slides

```
// Run all checks (overlap, bounds, empty_text, tiny_shapes)
verify_slides(slideIndex: 0)

// Run specific checks only
verify_slides(slideIndex: 0, checks: ["overlap", "bounds"])

// Returns: { slideIndex, shapeCount, issueCount, issues: [{ check, severity, shapes, message }] }
```

### Full Verification Loop

After completing work, verify ALL modified slides:

**Note on intentional overlaps:** Card layouts (TextBoxes + icons inside RoundedRectangles) and full-width decorative lines will always produce overlap warnings in `verify_slides`. These are expected — only act on overlaps between shapes that shouldn't be layered, or on overflow (shapes going off-slide). For large decks, run structural `verify_slides` on all slides but only visually verify the 4-5 most complex ones via subagent.

1. **Auto-size first:** If you edited text, set `autoSizeSetting = "AutoSizeShapeToFitText"` on those shapes — otherwise `verify_slides` sees stale dimensions:

```javascript
var shape = slides.items[0].shapes.getItem("28");
shape.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";
await context.sync();
```

2. **Structural check:** `verify_slides(slideIndex)` — overlap, bounds, empty text, tiny shapes
3. **Visual check:** Spawn a subagent for independent visual review. The subagent has no conversation context, providing an objective check. Use this prompt (replace N with slide index):

> Call screenshot_slide(slideIndex: N) to capture the slide, then review it for: text overflow or truncation, overlapping shapes or text, unreadable text (too small, poor contrast), misalignment or inconsistent spacing, empty or unused space, inconsistent styling (mixed fonts, colors, sizes). Return a JSON array of issues found, each with: severity (error/warning/info), category, description, and suggestion. If no issues found, return [].

4. **Fix and re-verify** until clean.

If overlaps/overflow found: shorten text, reduce font, reposition body content (not title), or split across slides.

Rules for visual review: never mention "the reviewer" to user — speak in first person ("I noticed..." not "The reviewer found..."). Only use for checking completed work, not initial inspection.

## Reading Content

```javascript
var slide = slides.items[0];
slide.shapes.load("items");
await context.sync();
var texts = [];
for (var i = 0; i < slide.shapes.items.length; i++) {
  var s = slide.shapes.items[i];
  try {
    s.textFrame.load("textRange");
    await context.sync();
    texts.push({ name: s.name, text: s.textFrame.textRange.text });
  } catch (e) { /* no text frame */ }
}
return texts;
```

## Custom Properties (API 1.7+)

```javascript
context.presentation.properties.custom.add("status", "draft");
await context.sync();
```

## Slide Backgrounds

Three distinct background-setting patterns — choose based on scope:

### Pattern 1: Single Slide Background
Sets background on one individual slide only. Use for slides that need a unique background.

```javascript
// NO # prefix — bare hex
slide.background.fill.setSolidFill({ color: "1A1A1E" });
```

### Pattern 2: Master Background via OOXML
Sets background for all slides via the slide master. Use in `edit_slide_zip` for blank decks when establishing a theme. See the [Slide Master & Theming](#slide-master--theming) section above.

### Pattern 3: Layout Loop
Sets background on all layouts, propagating to any slide using those layouts. Use when you want all layouts to share a background but don't want to modify the master directly.

```javascript
var masters = context.presentation.slideMasters;
masters.load("items");
await context.sync();
var slideMaster = masters.items[masters.items.length - 1];
var layouts = slideMaster.layouts;
layouts.load("items");
await context.sync();
for (var i = 0; i < layouts.items.length; i++) {
  layouts.items[i].background.fill.setSolidFill({ color: "355834" });
}
await context.sync();
```

## Shape Layering (Z-Order)

```javascript
shape.setZOrder("BringToFront"); // or "BringForward", "SendBackward", "SendToBack"
```

## Icons

Workflow: `search_fluent_icons` → `insert_image`

```
search_fluent_icons(query: "warning", top: 5)
// Returns: [{ id, description, isMono, contentTier, searchScore, svgUrl }]

insert_image(source: result.svgUrl, sourceType: "url", slideIndex: 0, left: 100, top: 100, width: 48, height: 48, color: "#FF5733")
```

**Variants:** filled (`isMono: false`, e.g. `Icons_Dog`) = colorful. Mono (`isMono: true`, e.g. `Icons_Dog_M`) = clean line-art. Prefer mono (`_M`) variants for professional decks — they're cleaner and recolorable.

**Sizing:** 36-48pt inline next to text, 72pt default, 72-144pt decorative hero.

**Coloring:** Pass `color` hex to `insert_image`. Only works with SVG sources — errors on non-SVG. Do NOT use `shape.fill.setSolidColor()` — that sets shape background, not SVG paths.

**Parallel operations:** When placing icons on multiple cards, search for ALL icons in parallel (multiple `search_fluent_icons` calls at once), then insert ALL icons in parallel (multiple `insert_image` calls at once). This is significantly faster than sequential one-at-a-time operations.

**Retry with alternative keywords:** If `search_fluent_icons` returns no good matches, retry with synonyms or related concepts:
- Abstract concepts → concrete objects: "innovation" → "lightbulb", "opinionated" → "compass", "security" → "shield"
- Actions → objects: "assessment" → "clipboard", "collaboration" → "handshake", "engineering" → "wrench"
- Compound concepts → simpler: "no lock-in" → "unlock", "faster delivery" → "rocket"

**Fallbacks** (if search returns nothing): geometric shape (filled circle + symbolic shape on top), or circle with single character in textFrame. Never use emoji or Unicode symbols.

## Slide Master & Theming

Use `edit_slide_zip` (targeting the full PPTX structure) to edit slide masters and themes.

**When to edit the slide master:**
- Blank deck (default theme, no content) — MUST edit first to establish theme
- Custom-styled deck (default theme, has content) — do NOT edit; existing slides define the style
- Template/existing (non-default theme) — do NOT edit unless user explicitly confirms a redesign

**Key files:**
- `ppt/slideMasters/slideMaster1.xml` — master shapes, background, text styles
- `ppt/slideLayouts/slideLayout1.xml` through `slideLayoutN.xml` — per-layout overrides
- `ppt/theme/theme1.xml` — theme colors, fonts, effects

Always: read → parse (DOMParser) → modify → serialize (XMLSerializer) → write. Never string concatenation.

### Setting Theme Colors

In `ppt/theme/theme1.xml`, find `<a:clrScheme>` and update:
- `<a:dk1>` — primary text (must contrast lt1)
- `<a:lt1>` — primary background (must contrast dk1)
- `<a:dk2>`, `<a:lt2>` — secondary dark/light
- `<a:accent1>` through `<a:accent6>` — accent palette

```javascript
function setColor(doc, parent, tagName, hex) {
  var el = parent.getElementsByTagName(tagName)[0];
  if (!el) return;
  while (el.firstChild) el.removeChild(el.firstChild);
  var clr = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:srgbClr");
  clr.setAttribute("val", hex);
  el.appendChild(clr);
}
```

### Setting Theme Fonts

Find `<a:majorFont>` and `<a:minorFont>` inside `<a:fontScheme>`, update `<a:latin typeface="...">`. Choose distinctive pairs — avoid defaulting to Calibri for both.

**Recommended font pairs** (heading + body):
- Montserrat + Lora
- Raleway + Open Sans
- Playfair Display + Source Sans Pro

### Setting Master Background

`<p:bg>` must be first child of `<p:cSld>` (before `<p:spTree>`) — use `insertBefore`:

```javascript
var cSld = masterDoc.getElementsByTagName("p:cSld")[0];
var bg = cSld.getElementsByTagName("p:bg")[0];
if (bg) cSld.removeChild(bg);
var fragment = new DOMParser().parseFromString(
  '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:bgPr><a:solidFill><a:srgbClr val="F5F0EB"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>', "text/xml").documentElement;
var imported = masterDoc.importNode(fragment, true);
var spTree = cSld.getElementsByTagName("p:spTree")[0];
cSld.insertBefore(imported, spTree);
```

For a gradient background, replace `a:solidFill` with `a:gradFill`:

```xml
<p:bgPr>
  <a:gradFill>
    <a:gsLst>
      <a:gs pos="0"><a:srgbClr val="1A1A2E"/></a:gs>
      <a:gs pos="100000"><a:srgbClr val="16213E"/></a:gs>
    </a:gsLst>
    <a:lin ang="5400000" scaled="1"/>
  </a:gradFill>
  <a:effectLst/>
</p:bgPr>
```

`pos` is in thousandths of a percent (0 = start, 100000 = end). `ang` is in 60,000ths of a degree (5400000 = 90 degrees, top to bottom).

### Setting Default Text Colors

Modify existing `<p:txStyles>` in slideMaster1.xml (find and modify — do NOT add new):

```javascript
var txStyles = masterDoc.getElementsByTagName("p:txStyles")[0];
var titleStyle = txStyles.getElementsByTagName("p:titleStyle")[0];
var titleDefRPr = titleStyle.getElementsByTagName("a:defRPr")[0];
var titleFill = titleDefRPr.getElementsByTagName("a:solidFill")[0];
if (!titleFill) {
  titleFill = masterDoc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:solidFill");
  titleDefRPr.insertBefore(titleFill, titleDefRPr.firstChild);
}
while (titleFill.firstChild) titleFill.removeChild(titleFill.firstChild);
var titleClr = masterDoc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:srgbClr");
titleClr.setAttribute("val", "FFFFFF");
titleFill.appendChild(titleClr);
// Repeat for bodyStyle
```

Critical: preserve element ordering. Find and modify existing elements — never add duplicates.

### Adding Decorative Shapes to Master

Append `p:sp` elements to `p:spTree` in slideMaster1.xml. Parse with DOMParser, import with `masterDoc.importNode(fragment, true)`, then `spTree.appendChild(imported)`.

For blank decks, always add at least one branding or decorative shape: accent bar along the bottom or side edge, thin divider line separating a header area, or subtle geometric shape as a background accent.

### Theming Rules

- NEVER override font colors on individual slides — all text should inherit from the master/theme
- NEVER add recurring visual elements (backgrounds, accent lines, decorative shapes, branding) to individual slides — they belong on the master or layout only
- Palette diversity: don't default to dark backgrounds. Light, warm, pastel, earthy, vibrant, and muted palettes are all valid. Match the tone of the content.

## Units & Positioning

All values in **points** (1 pt = 1/72 inch).

| Conversion | Formula |
|---|---|
| Inches to points | inches * 72 |
| cm to points | cm / 2.54 * 72 |

**Always use actual slide dimensions** from `inspect_deck`/`inspect_slide` response (`slideWidth`, `slideHeight`).

Common slide sizes:
- Standard 16:9: **960 × 540 pt** (13.33 × 7.5 in)
- Standard 4:3: **960 × 720 pt** (13.33 × 10 in)
- Widescreen: **1440 × 810 pt** (20 × 11.25 in)

| Reference | Formula |
|---|---|
| Full width | `slideWidth` |
| Center X | `slideWidth / 2` |
| Center Y | `slideHeight / 2` |
| Typical margin | 36 pt (0.5 in) |
| Title area | top 36 pt, height ~72 pt |
| Content area | top 120 pt to `slideHeight - 36` pt |
