# Autonomous engineering loop

How this project builds itself: an Opus-class agent takes a story from the roadmap and lands it as a merged PR, with quality enforced by gates so the human reviews a verdict, not a diff.

The same discipline that makes *slides* verifiable makes the *codebase* verifiable: every story has a machine-checkable done-signal (a cited [verification method](VERIFICATION.md)). The loop's job is to reach that signal.

## The loop

```
1. SELECT   pick the lowest-id unblocked story from ROADMAP/epics (deps satisfied)
2. BRANCH   <type>/<epic>-<slug>   (feat|fix|refactor|test|docs)
3. PLAN     restate the story's goal + its cited V-methods as the acceptance test
4. RED      write the failing test first (V2 negative fixture / V6 / V12 / V10) — watch it fail
5. GREEN    implement the minimum to pass; keep the change surgical (touch only what the story needs)
6. VERIFY   run every cited V-method in cost order (V1 → V2/V3/V9 → V4/V7 → V6 → V8 → V5 if applicable)
7. SELF-REVIEW  fork a fresh-eyes reviewer subagent (code + diff, no author context) → must return no blocking finding
8. BUILD    npm run build; stage dist/index.cjs (V13)
9. PR       Conventional-Commit title; body = the gate report + one-paragraph summary
10. CI      wait for green; if red, back to step 5 (fix forward, never merge red)
11. MERGE   squash-merge; delete branch; mark story done; back to 1
```

Steps 4–7 are the quality core. **Red-green is mandatory**: the negative case must demonstrably fail before the fix, or the test proves nothing. **Self-review by a fresh agent** is mandatory because an author goes blind to their own work — the same reason slide review is forked.

## Gates that make human review minimal

The human is not asked to read the implementation. They're shown, in the PR body:

- the **story goal** and which V-methods gated it,
- the **gate report** (V1 suite result, V4 golden-deck report if quality-relevant, V5 e2e status, V8/self-review verdict),
- a **one-paragraph** change summary and a change manifest (files touched, tools added/changed).

If every cited gate is green, the default is merge. The human intervenes only on: an escape-hatch (`execute_officejs`) surface change, a security-relevant change (V12), a schema/breaking-API change, or a gate the agent had to mark `warn`/override with justification.

## Guardrails for autonomous execution

- **One story per PR.** Small, reviewable, revertable. If a story needs >~400 changed lines, split it.
- **Never merge red.** A failing cited gate blocks merge. Fix forward or revert; do not disable the gate to go green.
- **Never weaken a gate to pass.** Adding coverage requires the mutation check (a bad input must turn it red). Lowering a threshold is a config change that itself needs justification + V10.
- **Migration behind the suite.** IR/tool refactors land tool-by-tool; the full suite stays green after each step (per [ARCHITECTURE.md](ARCHITECTURE.md) migration discipline).
- **Escape hatch is sacred.** Do not remove `execute_officejs`; promote patterns into typed tools alongside it.
- **Blocked?** If a story can't reach its done-signal (e.g. needs live PowerPoint the runner lacks, or an upstream epic isn't done), stop and surface it — don't fake the gate or mark it done.
- **Live-dependent stories.** Stories whose only proof is V5 (real Office.js behavior) run on the self-hosted macOS runner or locally; if unavailable, land the offline-verifiable part (V4) and leave the V5 acceptance explicitly pending, flagged for the human.

## Roles (who runs what model)

- **Executor** (Opus-class): runs the loop for one story end to end.
- **Reviewer** (fresh fork, step 7): independent code review; no author context.
- **Visual judge** (vision subagent): only for stories that change slide output, as the V8 gate.
- **Orchestrator** (optional): assigns unblocked stories to parallel executors when the dependency graph allows, respecting epic ordering.

## Definition of done (every story)

1. All cited V-methods pass.
2. Fresh-eyes self-review returns no blocking finding.
3. `npm run check` + `npm run build`/dist parity green (V1, V13).
4. PR merged; roadmap story marked done; any newly-unblocked stories noted.

If any of these is not literally true, the story is not done — regardless of how complete the code looks.
