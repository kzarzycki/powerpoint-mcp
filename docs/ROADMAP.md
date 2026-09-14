# Roadmap

Six milestones, ~18 epics, sequenced so each builds on a verified foundation. Epic detail (stories + per-story verification) lives in [epics/](epics/). Every story cites [verification methods](VERIFICATION.md); execution follows the [autonomous loop](AGENT-LOOP.md).

The through-line: **make the codebase trustworthy → build the deterministic quality gate → enforce it in the agent loop → make quality unconstructable-to-violate → make it on-brand → make the system self-maintaining.** Quality (priority 1) is addressed first and deepened at every milestone; token efficiency (2) and richness (3) ride on the same IR.

**Release scope.** The first public relaunch is the **reliable bridge**: [E00](epics/E00.md) plus install and contributor readiness. M1–M5 stay the product direction, not launch prerequisites. Transport authentication and an add-in CSP are deferred while the product is loopback-only and single-user — see [CODE-HEALTH.md](CODE-HEALTH.md) CH17 and CH18.

## Milestones

| M | Theme | Outcome | Epics |
|---|-------|---------|-------|
| **M0** | Trustworthy foundation | Connection and command-outcome defects are fixed with live evidence; the security posture the release ships with is decided and documented; the security-critical server paths have tests. Nothing new is built on sand. | E00, E01 |
| **M1** | DeckModel IR + deterministic gate | One normalized deck representation, populated identically by a live and an **offline** extractor; the deterministic linter stack runs headless in CI over golden decks. The always-on quality gate exists. | E02, E03, E04, E05, E06 |
| **M2** | Enforced verify loop + efficiency | Every mutating tool is paired with a verifier; `quality_check` runs the tiered stack; the structured visual judge replaces free-text review; responses are token-lean. The loop *enforces* quality instead of asking for it. | E07, E08, E09, E10 |
| **M3** | Quality by construction | A `compose_slide` engine with a constraint kernel emits common slide shapes that *cannot* violate the invariants. Reliability rises; token cost per slide drops. | E11, E12 |
| **M4** | Brand system | Brand packs (tokens + exemplars + layout catalog + assets) ingested from corporate decks, applied to live decks, and enforced by an `off_brand` gate with a bounded editable surface. | E13, E14 |
| **M5** | Consolidation & self-maintenance | Read surface consolidated onto orthogonal axes; prose rules that became code are deleted from the skill (doc-drift lint); the autonomous engineering loop is tooled and repeatable. | E15, E16, E17 |

## Epic overview

| Epic | Title | Milestone | Depends on | Verifies with |
|------|-------|-----------|-----------|---------------|
| **E00** | Bridge foundation (connection, command outcome, release posture) | M0 | — | V1, V2, V6, V12, V3 |
| **E01** | Bridge hardening, auth & error taxonomy | M0 | E00 | V6, V12, V1, V5 |
| **E02** | DeckModel IR + dual (live/offline) extractors | M1 | E01 | V2, V10, V4, V1 |
| **E03** | Quality-profile config | M1 | E02 | V10, V1 |
| **E04** | Deterministic linter engine (geometry/contrast/palette/typography/overflow) | M1 | E02, E03 | V2, V9, V4 |
| **E05** | Renderer abstraction + offline driver | M1 | E02 | V3, V4 |
| **E06** | Golden-deck CI regression stage | M1 | E04, E05 | V4, V13 |
| **E07** | Office.js codegen library | M2 | E02 | V3, V1 |
| **E08** | Token-efficient response layer | M2 | E02 | V7, V1 |
| **E09** | `quality_check` tiered tool + `verify_deck` + mutate↔verify pairing | M2 | E04, E06, E07 | V11, V4, V5 |
| **E10** | Structured visual judge + iterate-until-pass loop | M2 | E05, E09 | V2, V8, V4 |
| **E11** | Slide IR + layout engine + constraint kernel | M3 | E04, E07 | V9, V3, V4 |
| **E12** | Rich block library + `compose_slide` tool | M3 | E11 | V3, V9, V5 |
| **E13** | Brand-pack format + ingestion pipeline | M4 | E02 | V2, V10, V3 |
| **E14** | Brand apply + branded authoring + `off_brand` gate | M4 | E13, E04, E12 | V2, V4, V5 |
| **E15** | Read/inspect surface consolidation | M5 | E02, E08 | V7, V1, V11 |
| **E16** | Skill rewrite + rule-source-of-truth migration | M5 | E09, E12, E14 | V11, V1 |
| **E17** | Autonomous engineering-loop tooling | M5 | E06, E09 | V11, V4 |

## Dependency graph

```
E00 ─▶ E01 ─▶ E02 ─┬─▶ E03 ─▶ E04 ─┬─▶ E06 ─▶ E09 ─▶ E10
                    │              │         ▲         │
                    ├─▶ E05 ───────┘         │         └─(M2 loop enforced)
                    ├─▶ E07 ─────────────────┘
                    ├─▶ E08 ─▶ E15
                    └─▶ E13
E04 + E07 ─▶ E11 ─▶ E12 ─┐
E13 + E04 + E12 ─▶ E14 ──┤
E09 + E12 + E14 ─▶ E16   │
E06 + E09 ─▶ E17         │
```

E02 (DeckModel IR) is the keystone — most of M1–M5 depends on it. It lands **tool-by-tool behind the existing suite** (never a sweep), so M0 stays shippable throughout.

## Sequencing rationale

- **M0 first, non-negotiable.** CH01 (executed-but-reported-failed → duplicate slides) and CH02/CH03 (phantom "Multiple presentations connected") make the tool unreliable in exactly the multi-step flows the roadmap depends on; the security holes (unauth WS/MCP) are release-blockers. Building the IR on top of a bridge that drops responses would poison every later gate.
- **M1 is the keystone.** The offline extractor is the single highest-value absorption from the design review: it turns the flaky, M365-gated e2e (the current only guard) into an always-on headless gate, and it's a pure, incrementally-testable refactor. Everything downstream consumes the IR.
- **M2 makes the gate binding.** A gate the agent can skip isn't a gate; `quality_check` + mutate↔verify pairing + the iterate-until-pass loop convert prose "MANDATORY" into enforcement, and the token layer keeps the loop affordable.
- **M3 is prevention over detection.** Once the gate is trusted, `compose_slide` makes the common defects unconstructable — cheaper and more reliable than build-then-fix.
- **M4 layers brand on the proven engine** — brand conformance is just another linter family + a bounded editable surface, not a new architecture.
- **M5 pays down the surface and closes the loop** — consolidating tools, deleting prose that became code, and tooling the engineering loop so the project keeps improving itself.

## Milestone exit criteria

- **M0:** every CH## in [CODE-HEALTH.md](CODE-HEALTH.md) closed or explicitly deferred with reason; [E00](epics/E00.md) release gates met; V6 + V12 green; e2e connection-resilience test passing.
- **M1:** `npm run gate` runs offline in CI and is red on a mutated rule (V4 mutation check); live and offline extractors produce equal IR on the golden decks.
- **M2:** registry test passes (no mutating tool lacks a verifier); `quality_check` converges a seeded-defect golden deck to pass within N iterations using the mock judge; free-text `/review-slide-visual` retired.
- **M3:** `compose_slide` block set passes property tests (V9) — every generated layout satisfies the invariants; e2e screenshot pass on the golden deck.
- **M4:** a corporate `.pptx` ingests to a valid brand pack; applying it and authoring a branded slide passes the `off_brand` gate; violation fixtures flip it red.
- **M5:** `SKILL.md` contains no rule that is enforced in code (doc-drift lint green); the engineering-loop tooling can take a story file to a merged PR unattended on a sample story.

## Session-feedback stories (2026-07)

Mined from 39 real session transcripts ([feedback/session-analysis-2026-07.md](feedback/session-analysis-2026-07.md)); each story extends an existing epic and cites [verification methods](VERIFICATION.md). The same analysis added **CH23/CH24** to [CODE-HEALTH.md](CODE-HEALTH.md) (E00 scope) and confirmed E01, E08, E09 and the E02/E04/E10 gate architecture against field evidence. Two re-prioritizations: **E16-FB1 lands before M2** (it patches today's dominant failure mode until E07-FB1 ships), and applied-theme resolution (CH15 + CH24) must land before any E13/E14 brand work.

| Story | Epic | What | Feedback | Verify |
|-------|------|------|----------|--------|
| **E01-FB1** | E01 | Add-in auto-reconnect + WS heartbeat + server-side session resume; agent never needs the user to reopen the task pane | F01 | V6, V5 |
| **E01-FB2** | E01 | `bridge_health` doctor tool that names the failing layer (server / add-in socket / deck); single-instance port-conflict detection; fail-loud TLS misconfig; disconnect errors say what actually broke | F01, F08 | V6, V12, V1 |
| **E07-FB1** | E07 | Typed shape lifecycle tools: `add_shape`/`add_textbox`, `set_shape_geometry` (move/resize), `delete_shapes` with post-state read-back + dry-run, `reorder_slides` — modeled on the edit_shape_paragraphs design | F02, F03 | V2, V3, V5 |
| **E08-FB1** | E08 | Compact/cost-tiered reads: `inspect_layouts` compact mode (issue #98), cost-annotated tool descriptions steering scan-first, `read_slide_zip` default paths filter + size cap | F06 | V7, V1 |
| **E08-FB2** | E08 | `screenshot_slide` `savePath` param + one-line text caption in the result + documented recommended width | F07 | V7, V1 |
| **E10-FB1** | E10 | Visual-judge agent definition ships with the bridge screenshot tools in allowed-tools (the working subagent pattern), replacing the PNG hand-off | F13 | V8 |
| **E12-FB1** | E12 | `insert_icon(id)` with server-side URL resolution against a validated index; icon preview grid for human selection | F10 | V2, V5 |
| **E13-FB1** | E13 | Deck-level OOXML write path (presentation.xml: hide, sections, theme) — or an enforced, documented can't-do list so agents stop rediscovering the ceiling | F04 | V2, V4 |
| **E15-FB1** | E15 | One addressing vocabulary: 1-indexed slide numbers everywhere (or dual echo), `slideIndex`/`slideRange`/`shapeId` unified, zod coercion of numeric strings | F11, F12 | V1, V11 |
| **E16-FB1** | E16 | Interim skill content (front-loaded, before M2): copy-paste Office.js patterns — shape creation, fills, load/sync ceremony, batching discipline, known-missing APIs | F09 | V11 |
