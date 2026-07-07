# powerpoint-mcp — product & engineering docs

Design, roadmap, and specs for turning powerpoint-mcp into a system where AI agents build the best possible branded PowerPoint slides, provably, with minimal human review.

> Status: **alpha**. No stable APIs — breaking rework is expected. These docs describe the *target* and the sequenced path to it, not the current shipped surface. Current behavior is documented in the root `README.md` and `skills/powerpoint-mcp/SKILL.md`.

## Read in this order

1. **[VISION.md](VISION.md)** — what we're building, the priority order (quality > efficiency > richness), the live-PowerPoint differentiator, and honest limits.
2. **[ARCHITECTURE.md](ARCHITECTURE.md)** — target system: the DeckModel IR, the tiered quality gate, the compose engine, the brand system, the hardened bridge — and the migration discipline from v0.6.
3. **[QUALITY-GATES.md](QUALITY-GATES.md)** — the automated quality-gate stack (deterministic linters → renderer → structured visual judge), the rubric, and how it minimizes human review.
4. **[ROADMAP.md](ROADMAP.md)** — six milestones, ~18 epics, dependency graph, sequencing rationale, exit criteria.
5. **[epics/](epics/)** — one file per epic: stories with per-story verification references.
6. **[VERIFICATION.md](VERIFICATION.md)** — the reusable verification methods (V1–V13) every story cites.
7. **[AGENT-LOOP.md](AGENT-LOOP.md)** — the autonomous engineering loop that executes the roadmap under gate enforcement.
8. **[CODE-HEALTH.md](CODE-HEALTH.md)** — confirmed-bug backlog from adversarial code review (feeds M0).
9. **[INSPIRATIONS.md](INSPIRATIONS.md)** — borrowed ideas mapped to epics, with sources and pitfalls avoided.

## How this was produced

Multi-agent analysis on Opus-class models: five subsystem mappers, a six-dimension code review with per-finding adversarial verification (a skeptic tried to refute each bug), four web-research streams on similar projects, and a four-proposal architecture design panel scored by two independent judges. The synthesis and these documents were authored from those results. The confirmed findings, the winning architecture (pragmatic-incremental + the highest-value absorptions from the other three proposals), and the inspiration digest all trace back to that run.

## For an autonomous executor starting work

Pick the lowest-id unblocked story from [ROADMAP.md](ROADMAP.md) → [epics/](epics/), then follow [AGENT-LOOP.md](AGENT-LOOP.md). A story's done-signal is its cited [verification methods](VERIFICATION.md) passing. Start at **M0 / E00** — the foundation must be trustworthy before anything is built on it.
