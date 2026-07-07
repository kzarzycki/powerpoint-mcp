# Target architecture

The transport stays; the intelligence layer is new. We do **not** rearchitect the process model — one Node process serving HTTP static + WebSocket bridge + MCP is a deliberate simplicity decision that already ships (v0.6). We invest everything in a **DeckModel IR** and a **tiered quality gate** that make autonomous, verifiable slide-building cheap.

## Current state (v0.6)

```
Claude/Codex ──MCP(stdio|http:3001)──> Node process ──WS──> Office.js add-in ──> live PowerPoint
                                        (HTTP static :8080)   (WKWebView taskpane)
```

- The add-in exposes **one** action, `executeCode`: the server ships a JS string, the add-in runs it inside `PowerPoint.run()` via `new AsyncFunction`, serialized through a queue.
- ~26 MCP tools are **server-side generators of Office.js code strings**. OOXML editing round-trips through the add-in: export the deck via the Common API → mutate the zip/XML server-side (`jszip`, `@xmldom`) → reimport via `insertSlidesFromBase64`.
- Quality is **prose** in `SKILL.md` + a structural `verify_slides` tool + a free-text `/review-slide-visual` subagent.

What's wrong with it (see [CODE-HEALTH.md](CODE-HEALTH.md) for the bug list): the protocol is untyped and unversioned (drift caught only at runtime); the agent hand-writes Office.js/OOXML per slide (token-costly, error-prone); quality rules aren't enforced by code; there's no brand primitive; the visual review returns prose, not a machine verdict; neither channel is authenticated.

## Target state

```
                          ┌─────────────────────────────────────────────┐
   Claude / Codex         │  MCP tool layer (semantic, mutate↔verify)    │
        │                 │  read · mutate · compose · verify · brand    │
        ▼                 └───────────────┬──────────────────────────────┘
   MCP (stdio|http)                        │ emits / consumes
        │                 ┌───────────────▼──────────────────────────────┐
        │                 │  DeckModel IR  (zod-validated normalized deck) │
        │                 └───────▲───────────────────────▲──────────────┘
        │            live extractor│                       │offline extractor
        ▼                 ┌────────┴────────┐     ┌────────┴─────────┐
   Node process ──WS──>   │ Office.js bridge │     │ jszip OOXML read │  (no live PPT needed)
   (hardened: typed        │ (getImage, exec) │     └──────────────────┘
    errors, heartbeat,     └────────┬─────────┘
    auth, reconnect)                ▼
                          live PowerPoint (fidelity ground truth)

   Quality gate (tiered, cheap→expensive), consumes DeckModel + render:
     deterministic linters → renderer → structured visual judge → one report
   Config: one quality-profile (thresholds, palette, fonts, rubric, severities)
```

### 1. DeckModel IR — the keystone

A single **zod-validated normalized representation** of a deck: slides; shapes with geometry in points; text runs with *resolved* font/size/color; fills; the resolved theme palette; layout name and placeholder roles. Two extractors populate the **identical** schema:

- **Live extractor** over the bridge (`scan_slide` + `inspect_slide` + theme) — runtime path, exact geometry from Office.js rects.
- **Offline extractor** over a `.pptx` via `jszip` (already a dependency) — resolves scheme-color + tint/shade + master/layout inheritance to concrete hex. **Runs headless in CI without a live PowerPoint.**

Every gate, every read tool, and every verifier consumes the IR — not ad-hoc per-tool JSON. This is what makes the deterministic quality layer fast, token-stable, and CI-runnable, and it's a pure, incrementally-testable refactor. **Color resolution is the known landmine** (scheme + tint + inheritance): get it subtly wrong and the contrast/palette gates flood the agent with false findings it learns to ignore — so it gets its own fixtures and the highest correctness bar (`extractThemeFromZip` already has a confirmed bug returning `sysClr` tokens instead of hex; the extractor supersedes it).

### 2. Tiered quality gate

Cheapest-first so expensive judges never run on a deck that fails a free check:

1. **Deterministic linters** — pure `DeckModel → Finding[]` functions: geometry (margins, bounds, overlap, alignment, gap uniformity), contrast (WCAG ratio from resolved fg/bg), palette conformance (colors ⊆ theme ± tolerance), typography (font whitelist, size floor, hierarchy), text overflow (live: text-range bbox; offline: conservative heuristic reconciled against a render). Every finding carries a rule id, severity, and a slide/shape anchor. `verify_slides` is the runtime face of this catalog.
2. **Renderer** — `renderSlide(deck, i) → PNG` behind one interface with a live driver (`getImageAsBase64`, already built) and an offline LibreOffice-headless driver (for CI regression).
3. **Structured visual judge** — screenshot + weighted rubric → `{per-criterion score, anchored findings, pass/fail}`. Runs **only after** deterministic gates pass, and only as an on-demand / pre-done gate (never in the hot per-edit loop). Has a deterministic **mock/cache** mode so CI is reproducible. Replaces the free-text `/review-slide-visual` skill with a machine verdict.

All thresholds, the brand palette, the font whitelist, rubric weights, and per-rule severities live in **one quality-profile config** (zod-validated). One knob drives unit tests, golden fixtures, and runtime gating together. See [QUALITY-GATES.md](QUALITY-GATES.md).

### 3. Quality by construction (compose engine)

For common slide shapes, don't build-then-check — make the defect unconstructable. A small **layout engine** takes a typed **Slide IR** (`{layout, slots}`, not coordinates) and emits OOXML through the existing `edit_slide_zip` path. A **constraint kernel** applies invariants (bounds, contrast, overlap, font floor, no layout-cover) *before* emit, auto-reflowing where it can and rejecting where it can't. Exposed as a `compose_slide` tool over a **rich block library** (title, bullets, card-grid, stat-panel, icon-list, two-column, quote, image, table, chart). This is the AutoPresent lesson: a small high-level vocabulary raises reliability far more than raw-API breadth. `execute_officejs` remains the escape hatch — recurring ad-hoc patterns get promoted into typed blocks over time, never removed.

### 4. Brand system

A **brand pack** is a first-class on-disk artifact: `tokens` (locked palette hexes, font pair, spacing) + **exemplar slides** (few-shot targets — without them, branded output regresses to generic master slides, the documented M365 Copilot failure) + a **layout catalog** with semantic slot metadata + logo/assets. Built by **ingesting** a corporate `.pptx` (reusing the theme/layout extractors + python-pptx for master assets). Applied to a live deck, consumed by branded authoring tools, and checked by an `off_brand` gate in the linter catalog. The **editable surface is bounded**: static logos/disclaimers stay in the master, charts/tables are declared off-limits, so the agent can't mangle brand assets. Note the Office.js constraint: there is **no `applyTheme` API** — brand application goes through `insertSlidesFromBase64` from a branded base or direct OOXML injection into the master.

### 5. Hardened bridge (not rewritten)

Keep the process model; fix the confirmed protocol bugs: a **structured error taxonomy** (Office.js error → typed, self-correctable message so the agent can retry autonomously), request-id correlation logging, a **WS heartbeat** (prune half-open sockets instead of waiting for a 30s timeout), **execution-aware timeouts** (the timer must reflect execution, not queue wait — today it causes executed-but-reported-failed mutations and duplicate slides on retry), **stable presentation identity** (survive reconnect; no phantom "Multiple presentations connected"), and **auth** on both channels (WS origin allowlist + shared secret; loopback token on `:3001`).

## Migration discipline

Alpha allows breaking rework, but churn is the top execution risk. Rules:

- The IR refactor lands **tool-by-tool behind the existing test suite**, never in one sweep. Old tool tests stay green through each migration.
- The compose engine and brand system are **additive** — they sit alongside the existing tools; nothing is deleted until its replacement passes V4 + V5.
- As each prose rule becomes a code-enforced check, **delete it from `SKILL.md`** (enforced by the doc-drift lint, [V11](VERIFICATION.md)) so the skill shrinks toward a schema reference + escape-hatch guidance.
- `dist/index.cjs` is committed and used by installs; every server change re-runs the build ([V13](VERIFICATION.md)).

## Data flow, per agent step

`agent → MCP tool → (compose engine | codegen snippet) → WS command (correlated id) → add-in eval → DeckModel-normalized + compacted result → agent runs paired verifier / quality gate → loop`. The loop is the product: small mutate, machine-checkable verify, repeat — driven autonomously per [AGENT-LOOP.md](AGENT-LOOP.md).
