# powerpoint-mcp — product & engineering docs

Design and specs for turning powerpoint-mcp into a system where AI agents build the best possible branded PowerPoint slides, provably, with minimal human review. Work items — milestones, epics, stories — live on the [project board](https://github.com/users/kzarzycki/projects/2), not in this directory.

> Status: **alpha**. No stable APIs — breaking rework is expected. These docs describe the *target* and the sequenced path to it, not the current shipped surface. Current behavior is documented in the root `README.md` and `skills/powerpoint-mcp/SKILL.md`.

## Read in this order

1. **[VISION.md](VISION.md)** — what we're building, the priority order (quality > efficiency > richness), the live-PowerPoint differentiator, and honest limits.
2. **[ARCHITECTURE.md](ARCHITECTURE.md)** — target system: the DeckModel IR, the tiered quality gate, the compose engine, the brand system, the hardened bridge — and the migration discipline from v0.6.
3. **[QUALITY-GATES.md](QUALITY-GATES.md)** — the automated quality-gate stack (deterministic linters → renderer → structured visual judge), the rubric, and how it minimizes human review.
4. **[VERIFICATION.md](VERIFICATION.md)** — the reusable verification methods (V1–V13) every story cites.
5. **[AGENT-LOOP.md](AGENT-LOOP.md)** — the autonomous engineering loop that executes the board under gate enforcement.
6. **[CODE-HEALTH.md](CODE-HEALTH.md)** — findings and remediation decisions from adversarial code review and session feedback.
7. **[INSPIRATIONS.md](INSPIRATIONS.md)** — borrowed ideas mapped to epics, with sources and pitfalls avoided.

## Where the work lives

| | |
|---|---|
| Board | [Project 2](https://github.com/users/kzarzycki/projects/2) — `Status`, `Phase`, `Kind`, `Session` |
| Milestones | [M0–M5](https://github.com/kzarzycki/powerpoint-mcp/milestones) — sequencing and exit criteria |
| Epics | [`type:epic`](https://github.com/kzarzycki/powerpoint-mcp/issues?q=is%3Aissue+label%3Atype%3Aepic) — one per epic, stories as sub-issues |
| Stories | [`type:story`](https://github.com/kzarzycki/powerpoint-mcp/issues?q=is%3Aissue+label%3Atype%3Astory) — one engineering-loop iteration each |

These files hold what stays true regardless of which story is in flight: the target system, the gate stack, the verification vocabulary, the findings ledger. Anything with a status belongs on the board.

## How this was produced

Multi-agent analysis on Opus-class models: five subsystem mappers, a six-dimension code review with per-finding adversarial verification (a skeptic tried to refute each bug), four web-research streams on similar projects, and a four-proposal architecture design panel scored by two independent judges. The synthesis and these documents were authored from those results. The confirmed findings, the winning architecture (pragmatic-incremental + the highest-value absorptions from the other three proposals), and the inspiration digest all trace back to that run.

## For an autonomous executor starting work

Start at **[E00](https://github.com/kzarzycki/powerpoint-mcp/issues/125)** (M0), which carries the reliable-bridge release scope and its stories. Later epics are placeholders — write the stories when the epic is picked up. Follow [AGENT-LOOP.md](AGENT-LOOP.md). A story is done when its acceptance list is literally true.
