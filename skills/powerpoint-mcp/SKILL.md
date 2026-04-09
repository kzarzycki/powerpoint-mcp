---
name: powerpoint-mcp
version: 0.5.0
description: "Manipulate live, open PowerPoint presentations on macOS via Office.js MCP server. Use when Claude needs to: (1) create, edit, or inspect slides in a running PowerPoint instance, (2) add shapes, text, tables, or formatting to live presentations, (3) capture visual slide screenshots, (4) enable/configure the PowerPoint MCP server in a project, (5) execute Office.js code against open presentations. Distinct from the pptx file-editing skill — this works on presentations currently open in PowerPoint."
---

# PowerPoint Live Editing

Edit live, open PowerPoint presentations through an MCP bridge. Changes appear in real-time.

```
Claude Code  ──MCP HTTP (localhost:3001)──>  Bridge Server  ──WS──>  PowerPoint Add-in  ──>  Live Presentation
```

## Setup

When asked to enable or configure PowerPoint MCP in a project — follow the [setup guide](references/setup.md).

## MCP Tools

| Tool | Purpose | Key Parameters |
|------|---------|---------------|
| `list_presentations` | Discover connected presentations | — |
| `inspect_deck` | Deck overview: slide list + theme (colors, fonts). Use as first call | `presentationId?` |
| `inspect_layouts` | All slide layouts with names, indices (for `slides.add`), placeholders. `usedOnly` for fast mode | `usedOnly?`, `presentationId?` |
| `add_slide` | Add a new slide with a specific layout at a given position. Returns placeholder shape IDs | `layoutName`, `position?`, `presentationId?` |
| `inspect_slide` | Detailed shape inspector (~80 tok/shape): text, positions, fills | `slideRange`, `presentationId?` |
| `scan_slide` | Lightweight shape scanner (~40 tok/shape): IDs, types, positions | `slideRange`, `presentationId?` |
| `screenshot_slide` | Capture one slide as PNG (~1000 tok) | `slideIndex`, `width?` (default 720), `presentationId?` |
| `preview_deck` | Batch overview of all/selected slides with thumbnails + text | `slideRange?`, `imageWidth?` (default 480), `includeImages?`, `presentationId?` |
| `copy_slides` | Copy slides between two open presentations (data stays server-side) | `sourceSlideIndex`, `sourcePresentationId`, `destinationPresentationId`, `formatting?`, `targetSlideId?` |
| `insert_image` | Insert image from file path, URL, or base64 data (data stays server-side for file/url) | `source`, `sourceType` (`file`/`url`/`base64`), `slideIndex?`, `left?`, `top?`, `width?`, `height?`, `presentationId?` |
| `get_local_copy` | Get a local .pptx file path (passthrough for local files, exports cloud files to temp) | `presentationId?` |
| `read_deck_text` | Lightweight text extractor: titles + body as plain strings (~20x smaller than inspect_slide) | `slideRange?`, `includeNotes?`, `presentationId?` |
| `read_shape_paragraphs` | Read raw OOXML `<a:p>` paragraphs from a shape (preserves formatting) | `slideIndex`, `shapeId`, `presentationId?` |
| `edit_shape_paragraphs` | Replace paragraph content with raw OOXML (preserves bodyPr/lstStyle) | `slideIndex`, `shapeId`, `xml`, `presentationId?` |
| `read_slide_xml` | Read full slide OOXML or a specific shape's XML | `slideIndex`, `shapeId?`, `presentationId?` |
| `edit_slide_xml` | Replace full slide XML or a specific shape's XML | `slideIndex`, `xml`, `shapeId?`, `presentationId?` |
| `read_slide_zip` | Read multiple files from exported slide zip (slide XML, rels, charts) | `slideIndex`, `paths?`, `presentationId?` |
| `edit_slide_zip` | Update multiple zip files and reimport (auto-registers Content_Types for charts) | `slideIndex`, `files`, `presentationId?` |
| `duplicate_slide` | Clone a slide within the same presentation | `slideIndex`, `insertAfter?`, `presentationId?` |
| `verify_slides` | Check for overlapping, out-of-bounds, empty-text, tiny shapes, or unused placeholders | `slideIndex`, `checks?`, `presentationId?` |
| `edit_slide_chart` | Create chart from structured data (generates all OOXML automatically) | `slideIndex`, `chartType`, `title`, `categories`, `series`, `position?`, `options?`, `presentationId?` |
| `search_text` | Grep for slides — search text across shapes, tables, and speaker notes with regex support | `query`, `slideRange?`, `caseSensitive?`, `regex?`, `context?` (`shape`/`slide`/`none`), `includeNotes?`, `presentationId?` |
| `format_shapes` | Batch-apply fill and font formatting to shapes in one call | `slideIndex`, `shapes: [{ id, fill?, font? }]`, `presentationId?` |
| `execute_officejs` | Run arbitrary Office.js code in the live presentation | `code`, `presentationId?` |

`presentationId` is required only when multiple presentations are connected. Get it from `list_presentations`.

All positioning values are in **points** (1 pt = 1/72 inch). **Always read `slideWidth` and `slideHeight` from `inspect_deck` or `inspect_slide` response** — never assume 960 × 540. Common sizes: 960×540 (standard 16:9), 1440×810 (widescreen), 960×720 (4:3).

**Slide numbering**: Users refer to slides by 1-based number ("slide 3"), but all tools use 0-based indices. When a user says "slide 3", use `slideIndex: 2`. When they say "after slide 2", use `insertAfter: 1`.

### Read Tool Selection — pick the lightest tool that gives you what you need

| Tool | Scope | ~Cost | When to use | When NOT to use |
|------|-------|-------|-------------|-----------------|
| `inspect_deck` | All slides | ~5 tok/slide | First call — learn deck structure, slide count, dimensions | Need shape details |
| `scan_slide` | 1+ slides | ~40 tok/shape | Need shape layout/positions for editing | Need text content or fills |
| `inspect_slide` | 1+ slides | ~80 tok/shape | Need full details (text, positions, fills) | Just need positions — use scan_slide |
| `screenshot_slide` | 1 slide | ~1000 tok | Visual verification after changes | Looping over all slides — use preview_deck |
| `preview_deck` text-only | All slides | ~35 tok/slide | Content audit, find which slide has what | Need shape positions or fills |
| `preview_deck` + images | All slides | ~900 tok/slide | Visual audit of entire deck | Just need one slide — use screenshot_slide |

**Example**: a 20-slide deck: `inspect_deck` ≈ 100 tokens, `preview_deck` text-only ≈ 700 tokens, `preview_deck` with images ≈ 18000 tokens. Always prefer the lightest option.

### Tool Return Values

Key return formats to know:

- **`scan_slide`** returns `{ slideWidth, slideHeight, slides: [{ slideIndex, slideId, shapes: [{ id, name, type, left, top, width, height }] }] }` — `id` is a stable numeric string (use this for read/edit tools); `name` is locale-dependent (never use as selector); `type` is one of "GeometricShape", "TextBox", "Table", "Chart", "Picture", "Group"
- **`verify_slides`** returns `{ slideIndex, issues: [{ type, description, shapeIds }] }` — `type` is "overlap", "bounds", "empty_text", "tiny_shapes", "unused_placeholder", or "background_cover"; `shapeIds` are stable IDs
- **`search_fluent_icons`** returns `[{ id, description, isMono, contentTier, searchScore, svgUrl }]` — `isMono: false` = filled/colorful, `isMono: true` = outline/mono; pick highest `searchScore` matching intent; use `svgUrl` with `insert_image` (sourceType: "url") to insert
- **`read_shape_paragraphs`** returns raw OOXML `<a:p>` paragraph elements (does NOT include `<a:bodyPr>` or `<a:lstStyle>`)
- **`read_slide_zip`** returns `{ zipContents: { path: content }, allPaths: [...] }`

### Tool Behavior Notes

| Tool | Key non-obvious behavior |
|------|--------------------------|
| `edit_shape_paragraphs` | The `xml` field takes raw OOXML paragraph XML, not executable code. Preserves `<a:bodyPr>` and `<a:lstStyle>` automatically. Must auto-size shapes after edit. |
| `edit_slide_xml` | Two modes: `xml` (finished XML string) or `code` (JS that manipulates pre-parsed DOM server-side). Code mode preserves untouched attributes. Exported slide is ALWAYS `ppt/slides/slide1.xml` in the zip regardless of `slideIndex`. |
| `format_shapes` | Uses `getTextFrameOrNullObject()` internally. Cannot set corner radius, borders, or gradients — use `edit_slide_xml` code mode for those. Color format: hex without `#`. |
| `verify_slides` | Must auto-size shapes first or stale dimensions cause missed overlaps. Table overflow needs API fix, not OOXML. |
| `insert_image` | `color` param recolors SVG images only (e.g. Fluent UI icons from `search_fluent_icons`). Errors on non-SVG. `#` prefix required: `"#FF5733"`. |
| `execute_officejs` | Loaded values are snapshots — don't branch on stale reads after writes without re-load + re-sync. |

### OOXML sz Units (hundredths of a point)

| sz value | Point size | Use |
|----------|-----------|-----|
| `1400` | 14pt | Body minimum |
| `1600` | 16pt | Preferred body |
| `2000` | 20pt | Subheading |
| `2800` | 28pt | Section header |
| `3600` | 36pt | Slide title |
| `4400` | 44pt | Large title |

## Deck Type Detection

Before editing, determine the deck type. This determines the entire approach.

### Case 1: Blank Deck
**Detection:** `inspect_deck` shows only default slides, `inspect_slide` shows no custom content or colors.

Use `edit_slide_zip` FIRST to set up a complete theme before adding any slides. Do ALL of the following in a single `edit_slide_zip` call (targeting the slide master PPTX structure):
1. **Theme colors** — set the full `a:clrScheme`: dk1, dk2, lt1, lt2, and all six accents. Pick a cohesive palette suited to the topic and audience.
2. **Theme fonts** — choose a heading font (`a:majorFont`) and body font (`a:minorFont`) that pair well. Avoid Calibri for both.
3. **Master background** — set `p:bg` on the slide master.
4. **Default text colors** — update `p:txStyles` (title and body default text) so text contrasts the background. NEVER override font colors on individual slides.
5. **Decorative elements** — add at least one branding or decorative shape to the master (accent bar, divider line, subtle shape).

**Palette diversity rule:** Do NOT default to dark backgrounds. Light, warm, pastel, earthy, vibrant, and muted palettes are all valid choices. Match the tone of the content.

### Case 2: Custom-Styled Deck (Default Master)
**Detection:** `inspect_deck` shows default theme but `inspect_slide` reveals custom colors, fonts, and shapes on existing slides.

Do NOT create or modify the slide master. The existing slides ARE the design system.
- Before adding new slides, READ existing slides to extract visual style: background colors, font names, sizes, text colors, accent colors, shape styles.
- Pick the most representative slide as your style reference. Match its look exactly.
- Apply colors and fonts explicitly per-slide to match existing slides, since the master has no custom styles to inherit from.

### Case 3: Template or Existing Presentation
**Detection:** `inspect_deck` shows a non-default theme.

Default to PRESERVING the existing theme. New slides and additions should blend with existing colors, fonts, and layouts.

If the user requests a restyle or redesign — STOP before making edits:
1. Briefly describe the current template (master name, what it looks like).
2. Ask whether to (a) keep the current template and polish content within it, or (b) replace it with a new design.
3. Wait for the user's answer before proceeding.

## Planning & Elicitation

### When to ask clarifying questions BEFORE starting

For complex tasks (multi-slide decks, redesigns, data-heavy presentations), ask for missing context BEFORE doing any work. Do NOT assume details the user has not provided.

Triggers:
- "Make me a presentation about X" -- ask: audience, slide count, tone (formal/conversational), key points to cover
- "Turn this into slides" -- ask: structure (one topic per slide / grouped by theme), what to visualize vs. bullet-point, level of detail
- "Make this look better" / "Redesign" -- ask: focus (visual consistency / reducing density / restructuring flow), keep current structure or reorganize

### When NOT to ask

Simple, unambiguous requests ("Add a title slide", "Change the font to Arial", "Move this chart to slide 3") -- just do it. Factual or how-to questions -- just answer.

### Storyline-first rule (mandatory for multi-slide decks)

For multi-slide decks, PROPOSE THE STORYLINE FIRST -- slide titles and key points -- and get approval BEFORE creating any slides. Do NOT build 10+ slides without the user confirming the narrative arc.

### Layout prototype rule

When creating multiple slides that share a layout (e.g., one slide per team member, one per product), build ONE example slide first. Show it, get feedback on the design, then replicate across the remaining slides.

### Milestone checkpoints

For multi-step work, check in at key milestones. Show interim outputs and confirm before moving on. Do NOT build end-to-end without feedback.

## Workflow

1. **Discover**: `list_presentations` — find connected presentations
2. **Audit**: Check existing state — slide count, available layouts, which slides already have content. Use `preview_deck` for a visual overview, or `inspect_deck` then `inspect_slide` per slide. This is essential for resuming partial builds or modifying existing decks.
3. **Find**: `search_text` — grep for slides. Searches shapes, tables, and speaker notes. Use `context: "none"` for just slide indices, `"shape"` (default) for matching shapes, or `"slide"` for full slide context with all shapes. Supports regex.
4. **Detect deck type**: Determine blank / custom-styled / template (see above) — this decides whether to apply a theme first.
4. **See**: `screenshot_slide` — visually inspect specific slides
5. **Modify**: `execute_officejs` — build entire slides in a single call (all shapes, text, connectors, accents at once) for efficiency and to avoid mid-build visual flashing
6. **Verify**: full verification loop — `verify_slides` + `/review-slide` (see below). **Both steps are mandatory. Never skip `/review-slide`.**

Always inspect before modifying. Always verify after modifying. Every modified slide must pass both structural AND visual review before you move on or declare done.

### Incremental Deck Creation

For 3+ slides: build one slide at a time. Announce each slide before creating it ("Creating the Market Analysis slide..."). Use separate `execute_officejs` calls per slide -- this allows user feedback between slides.

Recommended flow for a multi-slide deck:
1. `edit_slide_zip` -- define theme, colors, fonts, background, decorative shapes (via slide master PPTX structure)
2. Title slide -- add slide with "Title Slide" layout, fill placeholders
3. Content slides -- one at a time, each in its own `execute_officejs` call
4. `verify_slides` -- check all slides for overlaps and unused placeholders

Do NOT build an entire multi-slide deck in a single call.

### Verification Loop

**MANDATORY** — run this after EVERY slide edit, including fixes. No exceptions.

1. **Auto-size first**: set `autoSizeSetting = "AutoSizeShapeToFitText"` on edited text shapes via `execute_officejs` — otherwise `verify_slides` sees stale dimensions
2. **Structural check**: `verify_slides` — overlap, bounds, empty text, tiny shapes, unused placeholders
3. **Text contrast check**: verify font color (set in the master's `p:txStyles`) contrasts the slide background. Flag and fix any per-shape color override that reduces legibility.
4. **Visual review**: invoke `/review-slide N presentationId` — this is NOT optional. The independent reviewer catches issues you cannot see from data alone (spacing, alignment, visual weight, contrast).
5. **Fix issues** and re-run from step 1. Repeat until only minor issues remain or only deliberate, acknowledged inconsistencies are left. Do NOT stop after one cycle if the reviewer flags real problems.

Do NOT declare success until the verify → fix → re-verify loop converges. Skipping `/review-slide` means you have NOT verified.

If overlaps/overflow: shorten text, reduce font, reposition body content (not title), or split across slides.

**Layout background rule**: NEVER create full-bleed rectangles to cover a layout-provided background. The layout background is part of a design system — it determines text colors, placeholder styling, and chrome visibility (logo, breadcrumb, dividers). Covering it destroys all of this and forces manual recreation. If a slide has the wrong background, switch to a layout that has the correct one.

**Manual checklist** (verify before declaring done):
- All placeholder shapes either filled with content or deleted
- Text color contrasts slide background on every shape
- No unused images or stale shapes from previous slide versions
- No shape text below 14pt
- All shapes have explicit `left` + `top` set
- No full-bleed rectangles covering layout backgrounds (enforced by `verify_slides` `background_cover` check — severity: error)

**Intentional overlaps**: When using card patterns (TextBoxes + icons inside RoundedRectangles), `verify_slides` will report many overlaps — these are expected by design. Also, decorative HR lines spanning the full width will overlap with adjacent elements. Only act on overlaps between shapes that should NOT be layered, or on overflow (shapes going off-slide).

**Efficient verification**: For large decks, visually verify only the most complex slides (high shape count, dense content) rather than every slide. Run `verify_slides` on all slides structurally, but pick 4-5 key slides for the visual subagent check.

### Visual Review — `/review-slide`

**When**: After every slide edit — step 4 of the verification loop. This is the final gate before declaring a slide done.

**How**: Invoke `/review-slide N presentationId` (N = 0-based slide index). Always pass the full presentationId — skips lookup and avoids ambiguity. The skill runs in a forked context with no conversation knowledge — it evaluates purely what it sees, eliminating confirmation bias.

**Why mandatory**: `verify_slides` catches structural issues (overlaps, bounds) but cannot detect spacing problems, visual imbalance, contrast issues, misaligned elements, or text overflow that Office.js doesn't report. Only a visual screenshot review catches these.

**Rules**:
- For large decks: run `verify_slides` structurally on all slides, but `/review-slide` only on the 4-5 most complex slides.

For `execute_officejs` code patterns, see [code-patterns.md](references/code-patterns.md).

## Tool Selection Hierarchy

Choose the right tool for each edit. Prefer higher tools in this table — fall back only when the preferred tool cannot express the change.

| Change type | Tool | Why |
|---|---|---|
| Fill color, font bold/italic/size/color/name | `format_shapes` | Declarative, 1 call per slide, no XML risk |
| Mixed formatting within one shape (some words bold, some not) | `edit_shape_paragraphs` | Office.js has no paragraph-level font API |
| Geometry (corners, borders), gradients, attributes Office.js cannot set | `edit_slide_xml` with `code` | DOM manipulation preserves untouched attributes |
| Complex layouts, new shapes, diagrams | `edit_slide_xml` with `code` | Full OOXML control |
| Simple text writes, shape creation, positioning | `execute_officejs` | Direct Office.js API |

**Key rule**: If `format_shapes` can express it, use `format_shapes`. Fall back to `edit_slide_xml` code mode only for properties Office.js cannot set (corner radius, borders, gradients, precise paragraph formatting).

**Never**: Use raw XML string replacement (`edit_slide_xml` with `xml` param) for formatting-only changes. OOXML is fully explicit — every omitted attribute is lost. Prefer `format_shapes` or `edit_slide_xml` with `code` (DOM manipulation preserves untouched attributes).

### Correction & Audit Workflow

When fixing style issues across multiple slides (e.g., after an audit finds inconsistent colors, fonts, or formatting):

1. **Process by slide, not by type** — apply ALL fixes for slide N in one call, then move to N+1. Do NOT sweep all slides for colors, then all slides for fonts, then all slides for corners.
2. **Reuse existing audit data** — if an audit file or previous tool output already describes the issues per slide, use it directly. Do NOT re-read slides with `inspect_slide` when you already have the information.
3. **Match tool to property**:
   - Office.js-expressible properties (fill, font bold/size/color/name) → `format_shapes` (one call per slide, all shapes)
   - OOXML-only properties (corners, borders) → `edit_slide_xml` with `code` (one call per slide, all shapes via DOM)
   - If a slide needs both, use two calls: `format_shapes` first, then `edit_slide_xml` code
4. **Verify after each slide**, not after each property type — `screenshot_slide` per slide to confirm all fixes applied correctly before moving on.

## User Preferences Persistence

Detect broad style preferences that apply across presentations and save them to memory:
- Save: "always use Oxford commas", "bold titles", "dark backgrounds", "keep slides evenly spaced", "always use 16pt body text"
- Do NOT save one-off, task-specific requests: "make this cell bold", "change this font to Arial", "move this chart to slide 3"

Before saving, check existing memory for duplicates. If the preference already exists, do not re-save it.

Use `/memory` or the project CLAUDE.md to persist preferences across sessions. Keep entries minimal -- one line per preference, grouped under a `## Slide Preferences` heading.

## Data Import Workflow

When the user provides data files (Excel, CSV, PDF) to populate slides:

1. Parse the file with `python3` (pandas, openpyxl, pdfplumber are available) to extract structured data
2. Use `execute_officejs` to populate slides with the extracted data

For .pptx template files: use `copy_slides` to import slides from another open presentation.

`insertSlidesFromBase64` rejects VBA macros, external references, OLE objects, and ActiveX controls. Only clean .pptx files pass through.

## OOXML Editing Workflow

**Prerequisite**: Load the `/pptx` skill for OOXML structure knowledge (namespaces, element anatomy, formatting rules).

For fine-grained formatting control beyond what Office.js properties expose, use the OOXML tools. See [ooxml-reference.md](references/ooxml-reference.md) for unit conversion, batching, and pipeline gotchas.

### Preferred: Code mode (`edit_slide_xml` with `code`)

One call — the code reads and modifies the pre-parsed DOM server-side. Only touched attributes change; everything else is preserved.

1. **Discover**: `inspect_slide(slideRange: "N")` → find shape IDs
2. **Edit + Write**: `edit_slide_xml` with `code` — DOM manipulation in one call
3. **Verify**: `screenshot_slide` — confirm visual result

**Sandbox context** (available in your code):

| Variable | Type | Purpose |
|---|---|---|
| `doc` | `Document` | Pre-parsed slide XML DOM |
| `findShapeById(id)` | `(string) → Element \| null` | Find `<p:sp>` by `<p:cNvPr id="...">` |
| `NS_P` | `string` | PresentationML namespace |
| `NS_A` | `string` | DrawingML namespace |
| `escapeXml(text)` | `(string) → string` | Escape `& < > " '` for XML |
| `serializeXml(node)` | `(Node) → string` | Serialize DOM node to XML string |
| `DOMParser` | constructor | Create new DOM documents for fragments |

**Example — remove rounded corners from multiple shapes:**
```js
edit_slide_xml(slideIndex: 0, code: `
  var ids = ["5", "7", "9"];
  for (var i = 0; i < ids.length; i++) {
    var shape = findShapeById(ids[i]);
    if (!shape) continue;
    var geom = shape.getElementsByTagNameNS(NS_A, "prstGeom")[0];
    if (!geom) continue;
    var avLst = geom.getElementsByTagNameNS(NS_A, "avLst")[0];
    if (avLst) while (avLst.firstChild) avLst.removeChild(avLst.firstChild);
  }
`, explanation: "Remove rounded corners")
```

**Example — change text color on a shape without losing formatting:**
```js
edit_slide_xml(slideIndex: 2, code: `
  var shape = findShapeById("3");
  var runs = shape.getElementsByTagNameNS(NS_A, "rPr");
  for (var i = 0; i < runs.length; i++) {
    var rPr = runs[i];
    var fills = rPr.getElementsByTagNameNS(NS_A, "solidFill");
    if (fills.length > 0) rPr.removeChild(fills[0]);
    var fill = doc.createElementNS(NS_A, "a:solidFill");
    var clr = doc.createElementNS(NS_A, "a:srgbClr");
    clr.setAttribute("val", "FFFFFF");
    fill.appendChild(clr);
    rPr.insertBefore(fill, rPr.firstChild);
  }
`, explanation: "Set text to white")
```

### Legacy: XML string mode

For single-shape paragraph edits, `read_shape_paragraphs` / `edit_shape_paragraphs` preserves `<a:bodyPr>` and `<a:lstStyle>` automatically. For full slide XML replacement, use `edit_slide_xml` with `xml` — but prefer code mode to avoid accidentally dropping attributes.

## Hard Limitations

Cannot do via Office.js — do not attempt:

- Insert images with precise shape-level control (use `insert_image` tool — positions via Common API, not shape API)
- Add animations or transitions
- Apply shadows or effects via Office.js (solid fills only via Office.js; gradients possible via OOXML `a:gradFill` in `edit_slide_zip`)

For charts, use `edit_slide_chart` (declarative) or `edit_slide_zip` (raw OOXML). Never approximate charts with geometric shapes. For slide masters/themes, use `edit_slide_zip` for full theme editing (colors, fonts, backgrounds, decorative shapes).

## Content & Design Rules

- Font minimum **14pt** everywhere, preferred body **16pt**
- Always explicitly set `font.size` — do not rely on defaults
- Max 3-4 key points per slide with short supporting text
- Prefer more slides with less content over fewer dense slides
- Use full slide area — stretch content to fill, don't leave large margins
- Never use emoji or Unicode symbols as icons — use geometric shapes as icon substitutes
- **Use icons to enhance content slides.** When a slide has key points, categories, or features, add relevant icons alongside the text. Always attempt `search_fluent_icons` before falling back to geometric shapes.

## Slide Layout Recipes

Common visual patterns for building slides. Adapt colors and content to the user's design system.

### Card Grid
RoundedRectangle as background → TextBox for title (offset ~85pt from left edge for icon space) → TextBox for body below title → Icon (36-48pt) at top-left corner of card.

Calculate card width: `(contentWidth - gaps) / numColumns`. Common configurations: 2x2, 3-across, 4-across, 5-across.

**Intentional overlaps**: Card patterns always report overlaps in `verify_slides` because TextBoxes and icons sit inside the RoundedRectangle. These are expected — only worry about overflow (shapes going off-slide) or unintended sibling overlaps.

### Icon + Text Blocks
Icon (36-48pt) left-aligned → Title TextBox at icon's right → Description TextBox below title, all inside a large RoundedRectangle container. Good for feature lists, "about us" sections, service descriptions.

### Key Numbers / Stats Panel
Large font number (accent color, 28-36pt) + small label below (14-16pt), stacked vertically with separator lines between entries. Good for KPIs, proof points, metrics panels.

### Pillar / Category Map
Vertical tall cards (equal width, evenly spaced) + horizontal bar spanning all pillars at bottom + dashed arrow connectors from each pillar down to the bar. Shows hierarchy: categories above → shared foundation below.

### Left-Right Content Split
Content panel (left, ~45% width) + stats/data panel (right, ~45% width) with a gap between. Good for combining narrative text with data points or proof points.

### Layered Stack
Horizontal rectangles stacked vertically with graduated fill color (darkest at top or bottom). Each layer has a title and description. Shows hierarchy, maturity levels, or technology stacks.

### Before/After Split
Two contrasting colored panels side by side (e.g., muted red for "without" vs green for "with"). Each panel lists bullet points. Optional full-width CTA bar below.

### Case Study / Reference Cards
3 equal-width tall cards, each with: header area (company/project name), description body, and metrics/outcomes section at the bottom.

### Cards with Tier/Tag Badges
Standard cards with a small colored RoundedRectangle "badge" overlaid (e.g., showing a tier level, category label, or status tag). Badge is typically 80-120pt wide, 20-28pt tall, positioned at top-right of the card.

## Gotchas

**XML:**
- Always escape `&` as `&amp;` in `<a:t>` — #1 cause of missing text
- OOXML is fully explicit — every omitted attribute is lost. **Prefer `format_shapes` for Office.js properties or `edit_slide_xml` with `code` for DOM manipulation. Both preserve untouched attributes. Avoid raw XML string replacement for formatting changes.**
- No `<!-- -->` comments in code strings — sandbox rejects with `SES_HTML_COMMENT_REJECTED`

**Office.js:**
- **Never use Office.js to read text content** — `textRange.text` returns plain text with all formatting stripped. Use `read_shape_paragraphs` for formatted content. Office.js is for shape metadata (IDs, positions, dimensions) and simple writes.
- Use `getTextFrameOrNullObject()` — never `.textFrame` directly (tables/images/charts throw)
- Loaded values are snapshots — don't branch on stale reads after writes (`hasText` stays stale after setting `textRange.text`)
- No `paragraphs` collection in PowerPoint Office.js
- Use `add_slide` tool to add slides with a layout at a given position (handles `slides.add()` + `moveTo` internally)
- Always use last master: `masters.items[masters.items.length - 1]` — earlier may be stale
- No `#` prefix for background colors: `{ color: "1A1A1E" }` not `"#1A1A1E"`
- Don't delete placeholders after writing text — `hasText` is stale, you'll delete what you just wrote. But DO delete genuinely unused placeholders (ones you never wrote to) — never leave empty placeholders on a finished slide.
- Shape IDs are stable and locale-independent. Shape names change with Office UI language. Always use ID.

**Charts:**
- Always register in `[Content_Types].xml`, include `<c:style val="2"/>`, don't hardcode series colors
- Stacked bars need `<c:overlap val="100"/>`, category axis `majorTickMark val="none"`

**Tables:**
- Height is auto-calculated — `shape.height` and OOXML `<a:ext cy>` are overridden
- Fix overflow via table API only: `cell.font.size` + `row.height`

## Working with python-pptx

For features Office.js cannot access (comments, chart data, embedded objects, master slides, custom XML parts), use `get_local_copy` to get a .pptx file path, then use python-pptx to process it.

- `get_local_copy` returns the existing file path for local files, or exports cloud files to a temp .pptx
- Reads the **saved** state — unsaved changes won't appear until the user saves
- Cached by revision number — only re-exports when the presentation has been saved since last export

## Error Handling

- **"No presentations connected"** — open PowerPoint with the add-in loaded
- **"Multiple presentations connected"** — specify `presentationId`
- **"Add-in disconnected"** — auto-reconnects; wait and retry
- **"Command timed out"** — simplify code or check PowerPoint responsiveness
- **Screenshot via execute_officejs overflows tokens** — always use `screenshot_slide` instead (returns MCP image block, not text)
