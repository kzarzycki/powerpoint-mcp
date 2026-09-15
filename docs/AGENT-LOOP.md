# Repository engineering loop

Every story runs in one isolated sibling worktree and has one exclusive remote claim.
The executable adapter is:

```sh
AGENT_SESSION=codex:<session-id> npm run loop -- claim --issue <url-or-number> --branch <slug>
```

`scripts/engineering-loop.ts` stores the claim and append-only phase history under
`refs/heads/loop-state/issue-<n>`. The claim is created on `origin` with create-only
semantics; a second session cannot overwrite it. The board and issue comments are
projections for people, not the lock. A failed claim writes a local rejected-claim
record and does not change the winning owner.

The claim command creates `../powerpoint-mcp--issue-<n>` from `origin/main` on
`loop/issue-<n>-<slug>`. Work never starts in the shared checkout. There is no
automatic stale-claim takeover: the owner or maintainer must resume or explicitly
park an abandoned claim.

## Fixed phases

1. **TRIAGED** — claim the issue and create its sibling worktree.
2. **SPEC** — write the issue-bounded spec, then run the standalone review:
   `npm run loop:review -- --cwd <worktree> --brief-file <brief> --output <verdict>`.
3. **SPEC_APPROVED** — approval is recorded with reviewer, findings and artifact hash.
4. **PLAN** → **PLAN_APPROVED** — write and independently review the ordered
   implementation plan.
5. **IMPLEMENTED** → **BRANCH_APPROVED** — implement, then review the complete diff
   from a fresh context. A rejected review leaves the phase in place for correction.
6. **GATES_GREEN** — run all required mechanical gates and record command, output and
   exit code.
7. **MERGED** — only after the story PR is green, the required evidence is attached,
   and the executing agent squashes the exact story branch.

Every review gate has three attempts. The attempt is reserved before the review runs,
so a crashed reviewer consumes an attempt rather than silently resetting the budget.
The third rejection changes the story to **PARKED**. No gate is weakened.
Rejections and later approvals remain in the append-only history.

## Project gates

The contributor runs these with Node 24 (`mise exec node@24.18.0 -- ...`):

- `npm run check` — Biome, TypeScript and the complete Vitest suite.
- `npm run build` followed by `git diff --exit-code dist/index.cjs` — committed bundle
  parity.
- `npm run test:e2e` for stories that exercise PowerPoint Web or the full MCP path.
- Live Office evidence for desktop/add-in stories: a disposable deck, cold launch,
  taskpane open, bridge `/health`, MCP `list_presentations`, taskpane close/reopen,
  and the same connection/presentation capture after recovery.

The fake Office host and `/health` endpoint do not prove Office.js behavior. Missing
PowerPoint, Microsoft 365 credentials, or a disposable deck is an explicit blocked
live gate, not a reason to substitute a weaker check.

## Review and landing

`engineering-review.ts` invokes the local `omp` CLI only. It does not import
Flowbench, Python, Omnigent or a daemon. It accepts only strict JSON:
`{"verdict":"APPROVE"|"REVISE","findings":[string]}`. Non-zero, timed-out,
empty or malformed reviewer output fails closed.

The executing agent may squash-merge ordinary loop PRs after the exact branch passes
the independent branch review, all required checks, CI and live evidence. Security
changes, `execute_officejs` changes, breaking APIs and gate exceptions still require
owner approval. Red CI is a failed gate; fix forward, retry within the cap, or park.

The issue comment and phase history are the human-facing record. `.loop/` files are
temporary worktree inputs and are not a second source of truth. A story is not done
until its PR is merged, its branch is removed, and every named acceptance criterion
has evidence.
