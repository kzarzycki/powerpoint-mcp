# Inspirations

What we borrowed and from where, mapped to the epic that carries it. Distilled from a survey of open-source slide MCPs, academic agentic-slide systems and their evaluation rubrics, commercial AI deck builders, and code-based deck ecosystems. Sources are cited inline; the full landscape reports are archived with the analysis run.

## Borrowed ideas → epics

| Idea | Source | Lands in |
|------|--------|----------|
| Render-and-judge visual gate (MLLM scores a rendered slide against a rubric, in an edit loop) — our live bridge captures **real** PowerPoint renders, an edge no python-pptx tool has | PPTAgent/PPTEval (arXiv 2501.03936); Anthropic pptx skill | E05, E10 |
| Judge as ~50 **atomic binary** checklist items scored fraction-passed, not a holistic 1–5 (Spearman 0.53 vs 0.30 — ~2× human agreement) | PresentBench (arXiv 2603.07244) | E10, QUALITY-GATES |
| Deterministic pre-VLM linters (WCAG contrast, overflow, off-canvas, overlap, font floor, margins, gaps) — VLMs are unreliable at geometry, so compute it exactly | VLM-SlideEval (arXiv 2510.22045); axe-core/Pa11y | E04 |
| Normalized deck IR both extractors emit; gates consume the IR, not live bridge | (synthesis; DeckModel from the design panel) | E02 |
| Offline OOXML extractor so the gate runs headless in CI without live PowerPoint | (design-panel absorption; jszip already a dep) | E02, E06 |
| One quality-profile config driving tests + fixtures + runtime gating | (design panel) | E03 |
| Small high-level action vocabulary (SlidesLib: 7 fns cut programs 170→13 lines, exec-rate 2%→80%) | AutoPresent/SlidesBench (arXiv 2501.00912) | E07, E12 |
| Layout = named container + named **slots**; agent emits `{layout, slots}`, template owns coordinates | Slidev, Marp, reveal.js | E11, E12 |
| Constraint-based layout: rules recomputed on every edit, "out of room" forces content cuts | Beautiful.ai Smart Slides | E11 |
| Brand kit = tokens + **exemplar slides** + layouts + assets (output regresses to generic without exemplars) | M365 Copilot; Gamma Brand Kit | E13 |
| Bounded AI-editable surface (static logos/disclaimers in master; charts off-limits) | Plus AI | E14 |
| Target-locking / activate-presentation to avoid editing the wrong open deck | ykuwai/ppt-mcp | E01 |
| Edit-from-reference over generate-from-blank | PPTAgent; Talk to Your Slides (arXiv 2505.11604) | E12, E14 |
| Fresh-eyes reviewer that must not grade its own output | Anthropic pptx skill | E10, AGENT-LOOP |
| Structured verdict (pass/fail + anchored findings) instead of prose review | (agent-surface map) | E10 |
| Change-scoped re-judging (only re-render changed slides — 50–80% CI cut) | Chromatic TurboSnap | E10 |
| Verifiable change manifest per edit | PPTArena (arXiv 2512.03042) | E09, AGENT-LOOP |
| Labeled-bounding-box render so the judge cites stable IDs + numeric corrections | Textual-to-Visual Self-Verification (arXiv 2502.15412) | E05 |
| Functional gate (cold reader answers MCQs from the render — catches pretty-but-empty) | Paper2Poster (arXiv 2505.21497) | E10 (optional) |
| Perceptual pixel-diff for template/brand regression (pixelmatch/odiff; "Layout" match level) | pixelmatch, odiff, Applitools | E06, E14 |
| Chart theme-color awareness; icon library with theme recolor + auto-fit | ykuwai/ppt-mcp; GongRzhe | E12 (charts), existing icon tool |
| ~30–40 semantic tools is the sweet spot; consolidate overlapping reads | GongRzhe; ykuwai (156-tool overhead) | E15 |

## Convergent evaluation rubric

Every academic system lands on the same three axes — use them for the visual judge:

- **Content** — accuracy, completeness, conciseness (not decorative-empty).
- **Design / visual** — layout, color, hierarchy, no-overlap. *(VLM strongest here.)*
- **Coherence** — cross-slide narrative flow. *(VLM's blind spot, ~0.55 correlation — back it with structural heuristics computed from DeckModel, not the VLM alone.)*

## Pitfalls others hit — designed around

- **Web/card paradigms lose fidelity on `.pptx` export** — Tome sunset its product over this. We stay PowerPoint-native. (→ VISION non-goals)
- **Holistic aesthetic scores are optimistic and weakly human-aligned** — decompose into binary checks. (→ E10)
- **Raw VLM aesthetic scores can't be trusted ungrounded** (informativeness/compositional bias) — pin the judge to a rubric requiring localized evidence; keep pure-aesthetic verdicts advisory. (→ E10, QUALITY-GATES warn-not-fail)
- **End-to-end image generation of slides is pretty but uneditable/garbled** — we stay programmatic. (→ VISION)
- **Raw-API-passthrough tools are agent-hostile** — typed, mid-granularity, semantic tools. (→ E07, E12, E15)
- **Brand output collapses to generic without exemplar slides** — gate on exemplar coverage before generation. (→ E13)
- **Constraint-engine rigidity feels templated** — keep `execute_officejs` and raw-OOXML escape hatches for bespoke slides. (→ VISION non-goals, E12)
- **The most feature-complete OSS reference (GongRzhe) is archived/unmaintained** — mine it for features, don't depend on it.
