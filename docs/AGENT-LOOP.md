# Repository engineering loop

Every story runs in one isolated sibling worktree and has one exclusive remote claim.
Start it with:

```sh
AGENT_SESSION=codex:<session-id> npm run loop -- claim --issue <url-or-number> --branch <slug>
```

The adapter stores append-only phase history under
`refs/heads/loop-state/issue-<n>`. Creation is remote and exclusive; a losing session
cannot overwrite the winner. The board and issue comments are human-facing
projections, not the lock. A losing claim leaves a local audit record and, when
publishing is enabled, an issue comment naming the winner.

The claim creates `../powerpoint-mcp--issue-<n>` from `origin/main` on
`loop/issue-<n>-<slug>`. Work never starts in the shared checkout.

## Fixed phases

1. **TRIAGED** — claim the issue and create its sibling worktree.
2. **SPEC** — write it with `loop spec --issue <n> --file <spec>`, then run the
   standalone reviewer: `npm run loop:review -- --cwd <worktree> --brief-file <brief>
   --output <verdict>`. Record that verdict with `loop review --gate spec --result
   <verdict>` before revising the artifact.
3. **SPEC_APPROVED** — the recorded reviewer verdict permits planning.
4. **PLAN** → **PLAN_APPROVED** — write and independently review the ordered plan.
5. **IMPLEMENTED** → **BRANCH_APPROVED** — run `loop implemented`, then review the
   complete branch from fresh context.
6. **GATES_GREEN** — run `loop checks --command '<gates>'`; command, output and exit
   code are recorded. `loop live --evidence <file>` records live evidence.
7. **MERGED** — after the exact story PR is green, run `loop merged --pr <url>
   --commit <sha>` for the squash merge.

Each review and mechanical gate has three attempts. A `REVISE` or failed gate counts
against the limit; the third failure parks the story. Reviewer output that is missing,
malformed or times out fails closed and does not become an approval. No gate is
weakened. Rejections and later approvals remain in the append-only history.

## Project gates

Run these with Node 24 (`mise exec node@24.18.0 -- ...`):

- `npm run check` — Biome, TypeScript and the complete Vitest suite.
- `npm run build` followed by `git diff --exit-code dist/index.cjs` — bundle parity.
- `npm run test:e2e` for PowerPoint Web or full-stack stories.
- Desktop/add-in stories require a disposable deck, cold PowerPoint launch, taskpane
  open, bridge `/health`, MCP `list_presentations`, taskpane close/reopen, and the
  same capture after recovery. Both observations must show `connections = 1` and
  exactly one connected presentation.

Red-green is mandatory: keep a negative case that fails before a bug fix. One story
stays in one PR. The `execute_officejs` escape hatch remains available. Migrations
stay behind the suite, and each story cites applicable methods from
`docs/VERIFICATION.md`. Never merge red, skip a cited gate, or lower a threshold to
make a gate pass.

Fake Office tests and `/health` do not prove Office.js behavior. Missing PowerPoint,
Microsoft 365 credentials, or a disposable deck is an explicit blocked live gate,
not a reason to substitute browser or fake-host output.

## Standalone review and landing

`engineering-review.ts` invokes local `omp` directly. It does not import Flowbench,
Python, Omnigent or a daemon. It accepts only strict JSON:
`{"verdict":"APPROVE"|"REVISE","findings":[string]}`. Non-zero, timed-out,
empty or malformed output fails closed. A verdict imported with `--result` is marked
as file-sourced in state history and must come from a separately run review.

The executing agent may squash-merge ordinary loop PRs only after independent branch
review, all required checks, CI and live evidence. Security changes, `execute_officejs`
changes, breaking APIs and gate exceptions still require owner approval. The issue
comment is best-effort; the remote state history is authoritative.
