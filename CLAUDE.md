# PowerPoint Office.js Bridge

## Project Goal

Build a system that lets Claude Code manipulate **live, open** PowerPoint presentations on macOS via Office.js APIs. First live-editing MCP solution for macOS — all others use python-pptx (file-based).

## Architecture

```
Claude Code  <--MCP STDIO/HTTP-->  Bridge Server (Node.js)  <--WS-->  PowerPoint Add-in (Office.js)
                                          |                                     |
                                    STDIO (default)                      WKWebView sandbox
                                    or HTTP (:3001/mcp)                  Office.js API 1.1-1.8
                                    localhost:8080 (HTTP)                executes commands on
                                    serves add-in files + WS             live presentation
```

Three components in one repo:

### 1. `addin/` - Office.js PowerPoint Add-in
- HTML/CSS/JS taskpane that loads inside PowerPoint
- Connects as WebSocket **client** to bridge server on load
- Receives JSON commands, executes Office.js API calls, returns results
- Manifest XML for sideloading

### 2. `server/` - Bridge Server (Node.js)
- HTTP server serving add-in static files (HTTPS opt-in via `BRIDGE_TLS=1`)
- WS (WebSocket) server for add-in connection (WSS when TLS enabled)
- MCP server (STDIO default, HTTP on port 3001 for standalone) exposing tools to Claude Code
- All three roles in one process for simplicity

### 3. `certs/` - Optional Local TLS Certificates
- Generated via `mkcert` for localhost (only needed when `BRIDGE_TLS=1`)
- `.gitignore`d

## Usage Documentation

For tool reference, code patterns, and usage — see the **powerpoint-mcp** skill at `skills/powerpoint-mcp/`.

## Technical Constraints

- **HTTPS/WSS opt-in** - Set `BRIDGE_TLS=1` to use HTTPS/WSS (requires certs via `npm run setup-certs`)
- **Add-in cannot host servers** - sandboxed in WKWebView, can only make outbound connections
- **Limited image API** - Image insertion via Common API `setSelectedDataAsync` (`insert_image` tool); no shape-level `addPicture()` yet (BETA only)
- **Charts via OOXML** - Office.js has no chart API; charts are created by injecting chart OOXML (`edit_slide_chart`)
- **No animations** - not exposed in stable APIs
- **Solid fills only** - no gradients, effects, or shadows
- **Points for positioning** - 1 point = 1/72 inch

## Work Tracking

Roadmap, epics and stories live on [GitHub project 2](https://github.com/users/kzarzycki/projects/2), linked to this repo. Nothing with a status lives in `docs/`.

- Hierarchy: milestone (`M0`–`M5`, exit criteria in the description) → epic (`type:epic` issue) → story/task/bug (sub-issues of the epic). Order between epics is `blocked by` issue dependencies, not prose.
- Board fields: `Kind` Epic/Story/Task/Bug · `Status` Backlog/Todo/In progress/In review/Done · `Phase` TRIAGED/SPEC/SPEC_APPROVED/PLAN/PLAN_APPROVED/IMPLEMENTED/BRANCH_APPROVED/GATES_GREEN/MERGED/PARKED · `Session` text. `Status` is the human column, `Phase` is the loop position — update both.
- Labels: `type:epic|story|task|bug`, `loop:needs-human`, `ready-for-agent`.
- Issue and milestone bodies follow the templates in `~/.agents/writing-style.md` → "Tickets and specs". They are written for the owner to accept; implementer facts go in the final "For the implementer" section.
- Option ids change whenever a field is edited — resolve them at call time. Recipes: the `github-project-board-setup` skill.
- `docs/` holds what stays true regardless of which story is in flight: `VISION.md`, `ARCHITECTURE.md`, `QUALITY-GATES.md`, `VERIFICATION.md`, `AGENT-LOOP.md`, and the findings ledger `CODE-HEALTH.md` (CHxx ids referenced from issues).

## Development Workflow

### Runtime
- Node `>=24` (`package.json` engines). An older Node still installs with an `EBADENGINE` warning, and local results can then differ from CI — switch to Node 24 before anything else. The repo has no version-manager pin file, so the check is `node --version`.

### MCP Dev Server
- `.mcp.json` uses `"type": "http"` pointing to `http://localhost:3001/mcp`; installed plugins start the server via the `SessionStart` hook, while source/dev runs still need a separately running server
- **Start**: `nohup node --experimental-strip-types ./server/index.ts --http --bridge > /tmp/powerpoint-mcp.log 2>&1 &`
- **Restart** (after code changes): `pkill -f "server/index.ts"; nohup node --experimental-strip-types ./server/index.ts --http --bridge > /tmp/powerpoint-mcp.log 2>&1 &`
- No build step needed for dev — runs directly from TypeScript source. Claude restarts the server itself; no user interaction needed.

### Gates
- `npm run check` = lint + typecheck + test. Run before every push.
- **After changing `server/` files**: `npm run build` and stage `dist/index.cjs` — it is committed and used by plugin/MCPB installs. The pre-commit hook enforces parity.
- CI runs `npm audit --omit=dev --audit-level=high`. A newly published advisory can fail `main` with no commit behind it. Fix by bumping the dependency, never by lowering the level. `dist/index.cjs` inlines the dependency tree, so a bump means rebuild + commit `dist/` too.
- Unit tests run against a fake Office host. They prove control flow, not PowerPoint behaviour. Anything host-dependent needs live evidence on the PR.
- Never weaken a test to go green. A test pinning incidental behaviour is deleted and replaced with a behaviour assertion.

### Branch & PR
- Feature branch `<type>/<short-description>` (e.g. `fix/tcc-sideload-prompt`, `feat/new-tool`).
- Commits and PR titles follow Conventional Commits (`feat:`, `fix:`, `chore:`, …). CI enforces this and auto-labels (`feat`→enhancement, `fix`→bug).
- After pushing, wait for CI to pass before calling the PR ready.
- Squash merge; delete the branch after merge.
- Note: `AGENTS.md` is a symlink to `CLAUDE.md` — edits land in `CLAUDE.md`.

### Code Quality
- TDD (red-green): failing test first, then implement.
- After implementation, `/simplify` the changed code.

## Key Technical Decisions

1. **Single Node.js process** for HTTP(S) + WS(S) + MCP (simplicity over microservices)
2. **TypeScript** for add-in and server (Office.js has good TS types)
3. **JSON command protocol** with request IDs for async response matching
4. **Plain HTTP/WS by default**, HTTPS/WSS opt-in via `BRIDGE_TLS=1` + mkcert certs
5. **Sideloading** for development (no Microsoft store submission needed)

## Command Protocol (WebSocket Messages)

```typescript
// Client (add-in) → Server
interface WSMessage {
  type: 'response' | 'error' | 'ready';
  id?: string;        // matches request ID
  data?: any;
}

// Server → Client (add-in)
interface WSCommand {
  type: 'command';
  id: string;         // unique request ID
  action: string;     // e.g. 'addShape', 'setText', 'getSlides'
  params: Record<string, any>;
}
```

## References

See `RESEARCH.md` for full research findings including:
- Detailed API capabilities per requirement set
- Code examples for all shape/text/table operations
- Existing solutions comparison
- macOS-specific issues and workarounds
- All relevant Microsoft documentation and GitHub issue links
