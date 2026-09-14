# Vision

Let AI coding agents (Claude Code, Codex, and others) build the **best possible** PowerPoint slides in a **live, open** presentation on macOS — on brand, and provably good, with the human reviewing as little as possible.

## Priorities, in order

1. **Output quality.** The rendered slide is what matters. Quality must be *enforced by the system*, not left to whether the agent remembered a prose rule.
2. **Token / time efficiency.** Every read, every edit, every verification step costs tokens and wall-clock. The cheapest tool that answers the question wins; the cheapest gate that catches a defect runs first.
3. **Content richness.** Charts, tables, icons, images, diagrams, dense-but-legible layouts — the deck should carry real information, not decorative emptiness.

When these conflict, the earlier one wins. We do not trade quality for a token saving; we do not add richness that fails a gate.

## The differentiator

Every other agent-facing PowerPoint tool is **file-based** (python-pptx, pptxgenjs): the model writes a `.pptx` it never sees rendered. We drive **live PowerPoint through Office.js**, so we can capture the *actual* rendered slide — real font fallback, real autosize, real overflow — and feed it back into a quality loop. No competitor has this. The roadmap is built around exploiting it.

## What "quality is enforced" means

Today quality lives in a 432-line skill file as prose the agent is asked to follow ("MANDATORY", "NOT optional" — language that exists precisely because nothing structurally enforces it). The target state moves each rule from prose into one of:

- a **deterministic linter** (geometry, contrast, typography, palette, overflow) computed exactly from slide data — cheap, always-on, CI-runnable;
- a **constraint kernel** that makes the defect impossible to construct in the first place (quality by construction);
- a **structured visual judge** (a fresh-eyes vision subagent scoring a rendered screenshot against a rubric) for the subjective residue the deterministic layer provably can't measure.

Prose that has become code gets deleted from the skill, so guidance and behavior can't drift apart.

## Honest limits (do not oversell the gates)

- **Visual quality is only partially machine-checkable.** Geometry, contrast, and overflow are exact; "does this read well" is not. The deterministic layer biases to **warn, not fail**, on subjective or intentionally-rich content so it never penalizes priority 3. Final aesthetic judgment stays with the visual judge and, for high-stakes decks, the human.
- **Render fidelity diverges.** Server-generated OOXML, LibreOffice-headless, PowerPoint Web, and PowerPoint macOS desktop all flow text and fall back fonts differently. **Live PowerPoint is the fidelity ground truth**; anything validated offline is an approximation that must be reconciled against a live render before "done".
- **Office.js API drift** is only caught by end-to-end tests against real PowerPoint, which are slow and gated on a Microsoft 365 account. Unit and offline layers cannot catch it.

## Non-goals

- No file-based / python-pptx fallback as the primary path — live editing is the entire point (python-pptx stays only as an escape hatch for parts Office.js cannot reach, e.g. reading embedded objects).
- No BETA/unsupported Office.js features (animations, gradients via API, shadows) chased before the quality loop is solid.
- No web/card layout paradigm (Gamma/Tome style) — it loses fidelity on `.pptx` export; Tome sunset its product over exactly this.
- No microservice split — one Node process (HTTP + WebSocket + MCP) is a deliberate simplicity choice.
- No big-bang rewrite — every change lands as a small increment behind a checkable signal (see [AGENT-LOOP.md](AGENT-LOOP.md)).

## How to read these docs

Start here, then [ARCHITECTURE.md](ARCHITECTURE.md) for the target system and [QUALITY-GATES.md](QUALITY-GATES.md) for the gate stack. [VERIFICATION.md](VERIFICATION.md) defines the reusable verification methods every story cites. [AGENT-LOOP.md](AGENT-LOOP.md) is the autonomous engineering loop that executes the board. [CODE-HEALTH.md](CODE-HEALTH.md) holds the findings and remediation decisions. [INSPIRATIONS.md](INSPIRATIONS.md) records what we borrowed and from where. The sequenced plan lives on the [project board](https://github.com/users/kzarzycki/projects/2).
