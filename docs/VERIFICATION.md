# Verification methods

Reusable verification recipes. Every roadmap story cites these by ID (e.g. "Verify: **V2** on the new check, **V4**") instead of re-describing how to check. An agent executing a story runs the cited methods and treats their pass as the done-signal.

Design rule: **prefer the cheapest method that can catch the defect.** Order of cost: V1/V10 (pure, ms) < V2/V3/V9 (fixtures, ms) < V4/V7/V11/V13 (offline gate, seconds) < V6 (simulated bridge) < V8 (vision subagent, one model call) < V5 (live PowerPoint, minutes, secret-gated).

| ID | Name | What the agent runs | Pass condition |
|----|------|---------------------|----------------|
| **V1** | Suite green | `npm run check` (biome + tsc + vitest) | Exit 0, no lint/type/test failure. Baseline for any change to `server/`. |
| **V2** | Check has both fixtures | Add a **positive** fixture (defect present) and a **negative/clean** fixture; run its unit test | Clean fixture → zero findings; seeded-defect fixture → exactly the expected finding with correct anchor + severity. A check without a negative fixture is not done. |
| **V3** | OOXML/codegen snapshot | Commit a snapshot of generated OOXML or Office.js code; parse the OOXML with `@xmldom` | Snapshot matches; generated XML is well-formed (parse throws on regression). Applies to chart/notes/slide builders. |
| **V4** | Offline gate on golden decks | `npm run gate` — offline extractor + deterministic linters + mock visual judge over committed fixture `.pptx` decks | Report equals the frozen snapshot. Mutation guard: flipping one rule threshold must turn a golden deck red (proves the gate is wired, not stubbed). |
| **V5** | Live round-trip e2e | Playwright vs PowerPoint Web (`npm run test:e2e`): apply the tool on the golden deck → `screenshot_slide`/`inspect_*` assert the expected structure | e2e job green. Gated on M365 secret; run locally or on the self-hosted macOS runner. **Only** method that catches Office.js API drift and real render behavior. |
| **V6** | Bridge resilience | Unit-simulate WS timeout, disconnect mid-command, double-`ready`, half-open socket | Pool ends in the correct state; in-flight requests reject with a typed error; no phantom pool entry; no silently-dropped late response. |
| **V7** | Token budget | Measure the tool's response payload on a sample deck against the committed baseline | Payload ≤ baseline (compaction did not regress); with field-selection, requested fields present and others absent; JSON still parses. |
| **V8** | Visual self-review (fresh eyes) | Fork the structured visual judge on the changed slide: it screenshots and returns `{verdict, findings[]}` with no conversation context | Zero `error`-severity findings. Used as the pre-done visual gate; the editing agent may **not** grade its own output. |
| **V9** | Property / invariant | Generate random valid inputs; assert the generated layout always satisfies invariants (in-bounds, no unintended overlap, font ≥ floor, contrast ≥ threshold) | All generated cases hold the invariants. Stronger than fixed fixtures for the compose engine. |
| **V10** | Schema validation | Feed the zod schema malformed and valid fixtures; change one config threshold and re-run a dependent check | Malformed rejected with a typed error, valid accepted; the threshold change provably changes a finding (config is live, not decorative). |
| **V11** | Consistency / no-drift | Registry test + doc-drift lint | Every mutating tool has a paired verifier; no rule that is enforced in code is also stated as a prose rule in `SKILL.md` (prevents guidance/behavior drift). |
| **V12** | Security assertion | Unit + request tests on the server | Path traversal rejected; `/health` carries no CORS header; WS upgrade rejects a foreign `Origin`; request body over cap is aborted pre-parse; logs contain no full session id or secret. |
| **V13** | Dist parity | `npm run build` then `git diff --exit-code dist/index.cjs` | Committed bundle matches source (plugin/MCPB installs use `dist/`; drift ships stale behavior). Enforced by the pre-commit hook and CI. |

## Notes for autonomous executors

- **Red-green first.** For a bugfix or a new check, write the failing V2/V6/V12 test *before* the fix and watch it fail, then make it pass. A story is not done until its negative case demonstrably fails without the change.
- **Offline before live.** Land and prove behavior with V4 (headless, always-on) before reaching for V5 (live, gated). Reserve V5 for what only real PowerPoint can confirm.
- **A gate that can't go red is not a gate.** Whenever you add V4/V10 coverage, include the mutation check that proves it fails on a bad input.
- **Warn-not-fail on subjective content.** Deterministic checks that touch rich/intentional design (dense card grids, decorative overlaps) emit `warn`, not `error`, so they never block a legitimately rich slide. Only objective failures (off-slide, sub-floor font, sub-threshold contrast) are `error`.
