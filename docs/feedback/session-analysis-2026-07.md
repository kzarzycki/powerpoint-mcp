# Session feedback analysis — 2026-07

What real usage says about powerpoint-mcp, mined from every Claude Code session that called the tools (19 transcripts, both `mcp__powerpoint-bridge__*` and `mcp__powerpoint-mcp__*` prefixes) plus every dev session on the two repos (26 transcripts, 5 overlap → 40 unique; 39 analyzed, 1 skipped: the mining session itself). One extraction agent per transcript, findings clustered across sessions, every quote machine-verified against the cited transcript line, and 6 of 14 issue clusters re-verified by independent adversarial agents (all 6 CONFIRMED — appendix A). Citations use local transcript paths under `~/.claude/projects/`; shorthands:

- `XEB` = `-Users-zarz-dev-personal-life-os-projects-Xebia-AI-Offering`
- `PPB` = `-Users-zarz-dev-powerpoint-bridge`
- `MCP` = `-Users-zarz-dev-agents-mcp-powerpoint-mcp`
- `INP` = `-Users-zarz-dev-personal-life-os-projects-inpost-pdlc-meeting`

## Top 10

Ranked by severity × frequency × impact on the stated goals (quality > token/time efficiency > richness). Roadmap column: existing epic/CH item it lands in, or the new story added to [ROADMAP.md](../ROADMAP.md) / [CODE-HEALTH.md](../CODE-HEALTH.md) in this change.

### 1. F01 — Bridge/add-in instability with no auto-recovery (bug, high, 8 sessions)

The dominant failure in real deck work: PowerPoint crashes during live use (5 crashes in one day, same StackHash, Office.js dispatcher co-triggered by add-in traffic), silent WS drops mid-batch, stale registrations reporting "No presentations connected" while the deck is visibly open. After any drop the only recovery is the user reopening the task pane and saying "bridge is back" — sessions stall on a human. The disconnect error text is misleading (tells you to open a file that is open) and nothing distinguishes bridge-down vs socket-dropped vs deck-not-open.

- PPB/9ab79f66-a0ff-4811-98d2-82bd0ea40165.jsonl:9 — "something is terribly wrong that on some use of MCP , powerpoint crashes. it crashed today N times."
- XEB/ffcbebe7-2032-47c2-ba52-3472580a52cc.jsonl:58 — "Add-in disconnected"

**Roadmap:** confirms **E01** and **CH01–CH04** as the right #1; adds stories **E01-FB1** (auto-reconnect + heartbeat + session resume) and **E01-FB2** (health tool naming the failing layer).

### 2. F04 — Single-slide export ceiling: presentation.xml edits silently dropped (bug + missing-capability, high, 5 sessions)

`edit_slide_zip` returns `success: true` for edits to files the reimport path cannot persist — `presentation.xml` is synthesized on export and regenerated on reimport (verified in `server/tools/xml.ts:141-144`). Hide-slide (`<p:sldId show="0">`), sections, theme/brand application, and version tracking are all deck-level and therefore unreachable; one session ended "This is useless. let's end here". Agents burn whole sessions rediscovering this ceiling.

- XEB/e6ad8624-759a-4709-a58b-0778fc925a09.jsonl:903 — "regenerates that file on every reimport, so any edit there gets dropped"
- XEB/e6ad8624-759a-4709-a58b-0778fc925a09.jsonl:859 — "still not hidden. explore it in subagent (and ask it to save to learnings why it didn't hide"

**Roadmap:** new **CH23** (error instead of fake success — E00) and story **E13-FB1** (deck-level OOXML write path or an enforced documented can't-do list; blocks the brand-templates goal).

### 3. F02 — No typed shape/geometry/slide-order tools; execute_officejs is the escape hatch (missing-capability, high, 10 sessions)

Every structural mutation — add/delete shape or textbox, move/resize, reorder slides — is hand-written raw Office.js. Agents repeatedly guess the API wrong (`addRectangle is not a function`, invented `Office.MSO_SHAPE_TYPE.RECTANGLE`, load-before-read retries — 4–5 distinct errors in one session before the working pattern), pay load/sync boilerplate on every geometry change (one session: all 15 execute_officejs calls were geometry), and hit a scale ceiling where large single calls stall.

- PPB/a738c16d-0592-4185-8e90-91459551cb93.jsonl:27 — "Error: slide.shapes.addRectangle is not a function."
- XEB/e6ad8624-759a-4709-a58b-0778fc925a09.jsonl:836 — "Probe what slide reorder APIs exist"

**Roadmap:** new story **E07-FB1** (typed add_shape/add_textbox, set_shape_geometry, delete_shapes, reorder_slides); long-term absorbed by **E11/E12** (`compose_slide`).

### 4. F03 — Unguarded raw deletes destroyed content and shipped undetected overlap (quality-gap, high, 2 sessions)

Raw `s.delete()` loops judged colleague WIP slides disposable and destroyed body text/notes the user wanted kept; a delete reported "5 icons deleted" while the old icons remained under new ones, and a 1920px screenshot review missed the doubling — the user caught it.

- XEB/6c5c6b0c-642e-4871-a612-15be4f7dbdec.jsonl:837 — "I see you vibed some other slides, which went a bit too quick. We need to be more dilligent in these moves."
- XEB/ed6aa6e1-b848-4ecd-9cea-e17d0856e53c.jsonl:550 — "heh, you don't see tht you haven't deleted the rpevious ones and now there is over another?"

**Roadmap:** confirms **E09** (mutate↔verify pairing); the delete tool in **E07-FB1** must read back post-state and support dry-run; **CH05** (reimport delete-without-rollback) is the same class.

### 5. F05 — Quality enforced only by prose; users hand-correct output (quality-gap, high, 3 sessions)

Nothing machine-checkable blocks "done" on a bad slide; the visual reviewer returns free text; deck-wide invariants (locked title geometry, uniform fonts) live only in the user's head, so users repeatedly took over cleanup themselves.

- MCP/e649244a-7adb-47bc-ace9-7b4f2b1ca090.jsonl:285 — "quality is enforced only by prose"
- XEB/127b9368-1ebc-47d9-adeb-2a5a737c46bf.jsonl:610 — "title shouldn't be moved, it should stay unmoved. Size of the font should also be same across slides."

**Roadmap:** validates the roadmap's center (**E02/E04/E09/E10**) with field evidence; deck-style invariants (locked header geometry/fonts) belong in the E04 linter family.

### 6. F09 — The shipped skill doesn't teach the Office.js patterns agents actually need (doc-skill-gap, high, 4 sessions)

With the skill loaded, agents still invented enums and missed the load/sync ceremony; one user said outright the skill gave the agent nothing and redirected it to an external office-js-docs MCP.

- PPB/a738c16d-0592-4185-8e90-91459551cb93.jsonl:105 — "i see the skill didn't give you anything. You have an mcp of officejs docs use it to learn how to do it"
- XEB/ed6aa6e1-b848-4ecd-9cea-e17d0856e53c/subagents/agent-a77dd2e553ee4a10e.jsonl:20 — "Error: undefined is not an object (evaluating 'Office.MSO_SHAPE_TYPE.RECTANGLE')"

**Roadmap:** new story **E16-FB1**, front-loaded: until E07-FB1 ships, the skill must carry copy-paste patterns (shape creation, fills, load/sync, batching, known-missing APIs). Cheapest quality lever available today.

### 7. F14 — Audit-surfaced defects still open: add_slide layout-by-name (#120), first-theme extraction (bug, medium, 3 sessions)

Most audit findings shipped as fixes (PRs #114–118), but `add_slide` still resolves layouts by name across all masters — multi-master decks (user: "not so rare in my company case") get the wrong layout and placeholder text lands in wrong shapes (issue #120, open). `extractThemeFromZip` takes the first `ppt/theme/*` file, which may not be the applied theme — wrong colors/fonts for any brand/contrast check built on it.

- MCP/d27b6225-ae7e-494a-bac7-9e325c027f3f.jsonl:1675 — "not so rare in my company case"
- PPB/437f46f2-82f0-45ff-ae29-a0b0d5ef4dce.jsonl:401 — "we haven't fixed this, yes I want it fixed in this release: server/index.ts:167 hardcodes MCP server version: '0.1.0'"

**Roadmap:** new **CH24** (#120, layout-by-id across masters — E00); extend **CH15** scope to applied-theme resolution. Both precede E13/E14 brand work.

### 8. F06 — Token-heavy inspection: inspect_layouts 10K+, inspect_slide ~80 tok/shape, read_slide_zip 1MB unfiltered (token-waste, medium, 6 sessions)

Listing layouts to pick one costs 10K+ tokens (issue #98 proposes a ~200-token compact mode); agents default to `inspect_slide` where `scan_slide` is half the cost; an unfiltered `read_slide_zip` dumped 1,039,104 chars and blew the token limit.

- PPB/d85e5b9f-21a2-4c3f-a569-f988b6af7bd9.jsonl:3 — "Currently it uses inspect layouts which can be easily 10K tokens or more (with default options)"
- PPB/a738c16d-0592-4185-8e90-91459551cb93.jsonl:163 — "i see we already have 100K tokens eaten. Can you tell me what is taking it?"

**Roadmap:** confirms **E08**; new story **E08-FB1** (compact layouts per #98, cost-annotated descriptions, read_slide_zip default filter + cap). **CH09** is the same theme.

### 9. F07 — Screenshot workflow inefficiency: no save-to-disk, image-only payloads, resolution churn, eyeball-tuning loops (token-waste, medium, 5 sessions)

`screenshot_slide` has no savePath (saving a PNG took a curl+base64 workaround); the result carries no text caption so vision-limited clients loop or misread it (qwen re-sent the same 56KB image 5×; minimax called a valid PNG "metadata"); the same slide got re-shot at 720→960→1920×3 (~1.7MB); geometry is tuned by guess-screenshot-adjust loops (15 screenshots in one session).

- PPB/013b20d2-1fa6-49ff-a5f0-9d5788c34253.jsonl:77 — "interesting, can you save the screenshot to disk and open it for me?"
- INP/6eea37db-be64-4219-8f31-8fe366fc917b.jsonl:28 — "can you see the screenshot ?"

**Roadmap:** new story **E08-FB2** (savePath + caption + recommended width); the eyeball loops are what **E10**'s structured judge and E04's geometry linters replace.

### 10. F13 — Reviewer subagents lack bridge MCP access; PNG hand-off workaround (friction, medium, 2 sessions)

The visual-reviewer agents couldn't reach the bridge, so the main agent rendered PNGs and re-dispatched — extra hops and long poll loops. Counter-evidence in the same corpus: subagents spawned with `mcp__powerpoint-bridge__*` in allowed-tools reached the bridge fine, so this is wiring, not architecture.

- XEB/94c2629f-acbc-4f3b-ab85-856723a7fc13.jsonl:2622 — "Reviewers don't have MCP access. Rendering slides to PNG via CLI, then re-dispatching."
- XEB/127b9368-1ebc-47d9-adeb-2a5a737c46bf.jsonl:388 — "Reviewer lacks bridge access. Rendering slide to PNG first."

**Roadmap:** new story **E10-FB1** (judge agent definition ships with the screenshot tools in allowed-tools). Prereq for the E10 iterate-until-pass loop.

## Additional findings

- **F08 — Setup/server-ops fragility** (friction, medium, 8 sessions): cache-clear silently unregisters the add-in; manifest served without CORS blocked Web sideload (found+fixed in-session); `BRIDGE_TLS=1` with missing certs left the bridge silently down; dev server and Desktop MCPB extension fight over port 8080 (issue #102). → **E01-FB2** doctor tool + fail-loud startup.
  - PPB/e93aa9e3-3f95-46b9-8dd3-72844f83ad69.jsonl:44 — "ah, I think we have 2 separate bridges started. How did it happen i don't know. but let's fix this."
- **F10 — Icon pipeline gaps** (medium, 3 sessions): `search_fluent_icons` is text-only (user built an HTML picker to choose icons); no insert-by-ID, hand-mapped `svgUrl`s 404. → new story **E12-FB1**.
  - XEB/ed6aa6e1-b848-4ecd-9cea-e17d0856e53c.jsonl:392 — "icons: i need to see the icons."
- **F11 — 0-indexed bridge vs 1-indexed PowerPoint UI** (medium, 2 sessions): users stop to double-check slide numbers; mandated 1-indexed convention mid-session. → **E15-FB1**.
  - XEB/6c5c6b0c-642e-4871-a612-15be4f7dbdec.jsonl:463 — "let's always use 1-indexed as in poweproint. it's more human oriented."
- **F12 — Inconsistent param names/types** (medium, 2 sessions): `slideRange` (string) vs `slideIndex` (number), singular vs plural `shapeId`; `screenshot_slide` hard-rejects `"0"`. Verified still present in `server/tools/inspect.ts` / `text.ts`. → **E15-FB1** (one addressing vocabulary + zod coercion).
  - XEB/6c5c6b0c-642e-4871-a612-15be4f7dbdec.jsonl:528 — "Invalid arguments for tool scan_slide"

## What works — do not break

- **F15**: the `read_shape_paragraphs` → `edit_shape_paragraphs` (+`edit_speaker_notes`) loop carried hundreds of edits across long sessions with ~zero failures and tiny payloads (57, 33 and 32 edits in single sessions). New tools should copy its shape: small typed params, paragraph granularity, tiny results.
  - XEB/94c2629f-acbc-4f3b-ab85-856723a7fc13.jsonl:2758 — "Yes — speaker notes pushed to all 3 slides via"
- **F16**: one `screenshot_slide` at width ~1600 is enough for a multimodal reviewer to produce fine-grained findings with zero retries; batched-build + screenshot-verify rebuilt whole layouts the user accepted. E10's machine verdict should formalize this loop, not replace it.
  - XEB/127b9368-1ebc-47d9-adeb-2a5a737c46bf/subagents/agent-abd5e95e6775025d6.jsonl:11 — "wraps to 4 lines while peers use 3 lines, breaking uniform card rhythm"
- **F17**: tiered reads (scan→inspect), dual text+image payloads (`preview_deck`), actionable error text (`list_presentations`), and `get_local_copy` as offline fallback all degraded gracefully — keep the ladder; extend the actionable-error style to the F01 disconnect cases.
  - MCP/e68c4199-868d-4459-acb5-56b933847ef4.jsonl:49 — "MCP server: **working**"

## Roadmap changes made in this PR

- **Confirmed, no change**: E01+CH01–CH04 as top priority (F01); E09 mutate↔verify (F03); E02/E04/E09/E10 gate architecture (F05); E08 token layer (F06).
- **Re-prioritized**: E16's skill content can't wait for M5 — the interim Office.js pattern story (E16-FB1) is front-loaded; theme-resolution correctness (CH15+CH24) must land before E13/E14 brand work.
- **New stories** in [ROADMAP.md § Session-feedback stories](../ROADMAP.md): E01-FB1, E01-FB2, E07-FB1, E08-FB1, E08-FB2, E10-FB1, E12-FB1, E13-FB1, E15-FB1, E16-FB1.
- **New code-health items** in [CODE-HEALTH.md](../CODE-HEALTH.md): CH23 (edit_slide_zip silent drop), CH24 (add_slide layout-by-name, #120).

## Appendix A — adversarial verification

6 of the 14 issue clusters sampled at random (python `random.Random(20260707)`, sample F01 F02 F04 F07 F12 F14); each handed to an independent agent instructed to refute using the primary transcripts and current source. All returned **CONFIRMED**:

| Cluster | Verdict | Note |
|---|---|---|
| F01 | CONFIRMED | crash partly upstream Office.js/Copilot — already framed as co-trigger |
| F02 | CONFIRMED | "6+ attempts" softened to 4–5 distinct errors; tool absence verified in source |
| F04 | CONFIRMED | silent success verified at `server/tools/xml.ts:141-144` |
| F07 | CONFIRMED | no savePath/caption verified at `server/tools/inspect.ts:359-437` |
| F12 | CONFIRMED | schema mismatch verified in `inspect.ts`/`text.ts` |
| F14 | CONFIRMED | #120 open via `gh`; first-theme grab verified in `xml-helpers.ts:348-360` |

All 47 citations in the underlying cluster analysis were additionally machine-checked (quote must appear at the cited transcript line): 47/47 pass.

## Appendix B — corpus

Discovery: `grep -rl '"name":"mcp__powerpoint' ~/.claude/projects --include='*.jsonl'` → 19 usage transcripts (all 19 analyzed), plus all transcripts under the three dev project dirs (26, of which 5 overlap usage) → 40 unique; 39 analyzed, 1 skipped (the mining session's own live transcript). 489 powerpoint tool calls, 38 tool errors in the corpus. Heaviest sessions: 2026-05-21 Xebia-AI-Offering cluster (106+87+71+37 calls — real deck builds), 2026-05-23 (24 calls / 9 errors) and 2026-04-13 (18 calls / 9 errors). Sessions with zero findings (trivial or unrelated): 13. Per-session extracts and the full cluster analysis were produced in the analysis workspace; this file is the durable summary.
