---
name: review-slide-visual
description: >
  Visually review a PowerPoint slide for layout issues, spacing problems, contrast errors,
  and alignment inconsistencies. Spawns an independent reviewer with no conversation context.
  Use after completing slide edits to verify visual quality — not for initial inspection.
  Trigger on "review the slide", "how does it look", "check the visual", "verify the layout",
  or proactively after slide editing is complete.
context: fork
agent: Explore
model: opus
allowed-tools: mcp__powerpoint-mcp__screenshot_slide, mcp__powerpoint-mcp__list_presentations
argument-hint: <slide-index> [presentation-hint]
---

## Resolve presentation

- If $ARGUMENTS[1] looks like a full presentationId (file path or generated ID), use it directly — skip list_presentations.
- Otherwise, call list_presentations first. Then:
  - If $ARGUMENTS[1] is a fuzzy hint (e.g., "ai enablement v3"), find the best match. If exactly one matches, use it. If ambiguous or no match, stop and report — do NOT guess.
  - If $ARGUMENTS[1] is not provided and only one presentation is connected, use it.
  - If $ARGUMENTS[1] is not provided and multiple are connected, stop and list them — do NOT guess.

## Capture and review

Call screenshot_slide with slideIndex set to $ARGUMENTS[0], width 1440, and the resolved presentationId.

Visually inspect this slide. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Content crammed at top of container with empty bottom half
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Peer elements with inconsistent sizing (cards that should be same height aren't)
- Low-contrast text (e.g., light gray on cream background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without contrasting circle)
- Icons/decorative elements too small relative to their container, or blending into background
- Text boxes too narrow causing excessive wrapping
- Lack of visual hierarchy (headings indistinguishable from body text)
- Leftover placeholder content

Report ALL issues found, including minor ones.
Be specific about locations (e.g., "bottom-right text box overflows the slide edge").
Quantify when possible (e.g., "bottom ~60% of the card is empty", "gap is ~2x wider than its neighbor").
Group issues under category headings: Layout & Spacing, Text & Content, Contrast & Visibility.
Number each issue. Use **bold label** — description format.
Do NOT suggest content/wording changes — only evaluate visual presentation.
