# Automated quality gates

The heart of the project: how a slide is proven good with minimal human review. Three tiers, cheapest first, all driven by one config.

## Principles

- **Tiered by cost.** Deterministic linters (free, exact) run first; the vision judge (a model call) runs only on decks that already passed them. A deck never spends a judge call to be told it has off-slide text.
- **Binary checklist over holistic score.** Judge output is decomposed into atomic yes/no items scored as fraction-passed — not a single 1–5 aesthetic rating. Research (PresentBench) shows this nearly doubles agreement with human judgment (Spearman 0.53 vs 0.30). It's also cheaper to evaluate and auditable: you can see *which* item failed.
- **Deterministic where possible, VLM only where necessary.** Vision models are provably unreliable at pixel geometry (VLM-SlideEval). Overlap, bounds, contrast, font size, margins, alignment are computed **exactly** from DeckModel geometry. The judge is reserved for the subjective residue: visual balance, hierarchy, "does it read well".
- **Warn, don't fail, on rich/subjective content.** Only objective failures are `error` (block done): off-slide, sub-floor font, sub-threshold contrast, empty placeholder, malformed OOXML. Dense card grids, intentional decorative overlaps, and stylistic choices emit `warn` so the gate never punishes priority-3 richness.
- **Live is ground truth.** Offline checks are an always-on approximation; before "done" on a real deck, reconcile against a live PowerPoint render.
- **A gate that can't go red isn't a gate.** Every gate ships with a mutation check ([V4](VERIFICATION.md)) proving a bad input turns it red.

## Tier 1 — Deterministic linters (`verify_slides` catalog)

Pure `DeckModel → Finding[]` functions. Each finding: `{ ruleId, severity: error|warn|info, slideIndex, shapeIds[], message, measured, threshold }`.

| Family | Rules | Severity default | Source of truth |
|--------|-------|------------------|-----------------|
| Geometry | off-slide / clipping, pairwise unintended overlap, edge margin < 0.5", inter-element gap < 0.3", column/row misalignment, uneven peer sizing | error (bounds), warn (spacing) | DeckModel rects (exact) |
| Contrast | WCAG ratio of resolved text vs resolved background/fill; AA 4.5:1 body, 3:1 large | error below threshold | resolved colors |
| Palette | colors ∉ theme/brand palette ± tolerance | warn (theme), error (brand pack) | resolved palette |
| Typography | font ∉ whitelist, size < floor (14pt), heading/body hierarchy inversion | error (floor), warn (whitelist) | DeckModel runs |
| Content | empty/placeholder text left in, tiny shapes, unused placeholder, background-cover rect over layout | error | DeckModel |
| Overflow | text exceeds its box, table-cell overflow, distorted image aspect | error | live bbox; offline heuristic + render reconcile |
| Structure (charts) | series length ≠ category length, non-finite values, malformed/mis-ordered chart OOXML | error | OOXML parse |

These run headless in CI over golden decks (via the offline extractor) and at runtime over the live deck. Existing `verify_slides` checks (overlap, bounds, empty_text, tiny_shapes, unused_placeholder, layout_drift, background_cover) are the seed of this catalog.

## Tier 2 — Renderer

`renderSlide(deck, i) → PNG` behind one interface:
- **Live driver**: `getImageAsBase64` (already built) — real PowerPoint pixels, the fidelity ground truth.
- **Offline driver**: LibreOffice headless (`soffice --convert-to`) — for CI golden-deck regression when no live PowerPoint is available. Known to diverge from PowerPoint on text-flow/font-fallback; used for *regression* (did this change move pixels?), not absolute fidelity.

Optionally overlays labeled bounding boxes with shape IDs on the render so the judge can cite stable IDs and return numeric corrections ("increase card height ~1.2×").

## Tier 3 — Structured visual judge

A **fresh-eyes** vision subagent (no conversation context — the editing agent must not grade its own work) receives the render(s) + the weighted rubric and returns structured JSON:

```
{ verdict: pass|fail,
  criteria: [{ id, dimension, passed: bool, evidence, shapeIds? }],
  score: <fraction of items passed> }
```

Rubric dimensions follow the convergent core across every academic system (PPTEval, PresentBench, Paper2Poster, DeepSlides):

- **Design / visual** — layout balance, alignment, hierarchy, whitespace, color harmony. *(VLM strongest here.)*
- **Content** — clarity, completeness, conciseness; not decorative-empty. *(Optionally hardened by a functional check: a cold reader answers questions from the render alone — catches pretty-but-empty.)*
- **Coherence** — cross-slide flow. *(VLM's blind spot, ~0.55 correlation — back it with structural heuristics, e.g. consistent layouts/fonts across slides computed from DeckModel, not the VLM alone.)*

Runs on-demand and as the pre-done gate, not per-edit. Deterministic **mock/cache** mode returns frozen verdicts so CI is reproducible and doesn't spend model calls.

## The `quality_check` tool and the iterate-until-pass loop

One MCP tool runs the tiered stack and returns a single prioritized report (deterministic findings first; judge only if they pass). The autonomous loop:

```
edit → quality_check
  ├ deterministic errors?  → return anchored findings → fix → recheck   (cheap, tight loop)
  ├ deterministic clean    → render + judge
  │     └ judge fail?       → return anchored findings → fix → re-judge
  └ all pass OR budget hit  → done (report attached to the change)
```

Change-scoped: only slides whose content/template/inputs changed are re-rendered and re-judged (the TurboSnap pattern — 50–80% less work). Bounded by a token/iteration budget so it always terminates.

## CI: golden-deck regression

A CI job runs the **offline** path (extractor → linters → offline render → mock judge) over committed fixture `.pptx` decks and snapshot-asserts against frozen reports ([V4](VERIFICATION.md)). Mutating any rule flips a fixture red. This makes the entire deterministic layer an always-on gate that external contributors and default CI can run without a Microsoft 365 account — closing the biggest verification gap (today CI runs only mocked unit tests; the live e2e that exercises real behavior never runs in CI).

## What this buys the human

The human reviews a **verdict and a short findings list**, not the deck. A passing `quality_check` report (deterministic clean + judge pass, attached to the PR/change) is the evidence. The human looks only when the gate is uncertain (warns, or judge score in a gray band) or the deck is high-stakes. That is the "minimum text to review" goal, delivered.
