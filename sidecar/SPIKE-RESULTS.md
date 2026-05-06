# ACP × PowerPoint MCP — Milestone-1 Spike Results

**Status:** PASS · 2026-05-06 · branch `acp-spike`

## What was validated

End-to-end wiring: Node sidecar spawns `@agentclientprotocol/claude-agent-acp` over stdio, exposes the ACP wire protocol over a WebSocket, intercepts `session/new` to inject the PowerPoint MCP server config, and the agent reaches the bridge's HTTP MCP endpoint and successfully invokes a tool.

```
test-client (ws) ──► sidecar (this dir) ──stdio──► claude-agent-acp ──HTTP──► powerpoint-bridge (:3001/mcp)
```

## Run

Bridge running on `:3001/mcp` (MCP) + `:18080` (taskpane bridge, no taskpane connected — irrelevant for this spike).

```
cd sidecar && node --experimental-strip-types test-client.ts
```

Result: `SPIKE PASS — agent invoked PowerPoint MCP` in ~13 s, $0.33.

Observed events on the wire:
1. `initialize` round-trip OK; agent advertised `mcpCapabilities: { http: true, sse: true }`.
2. `session/new` outbound was augmented by the sidecar — confirmed by `[sidecar:inject] session/new augmented: +mcpServers[powerpoint] +_meta.systemPrompt`.
3. Agent issued tool call: `mcp__powerpoint__list_presentations` (toolCallId `toolu_…`), kind `other`, status `pending`.
4. Tool response came back from the bridge: `"No presentations connected. Open a PowerPoint file with the bridge…"` — proves the HTTP MCP transport returned real data, not a wire stub.
5. Agent emitted `agent_message_chunk` deltas → `stopReason: end_turn`.

## Load-bearing facts confirmed

| Question | Answer |
|---|---|
| Wire format on stdio | newline-delimited JSON-RPC 2.0 (no Content-Length headers) |
| Method names | `initialize`, `session/new`, `session/prompt`, `session/update` (notification), `session/request_permission` (server→client request) |
| `mcpServers` of `type: "http"` | Works. Required fields: `name`, `url`, `headers: []` |
| System prompt injection | `params._meta.systemPrompt = { append, excludeDynamicSections }` — agent honored the append on first turn |
| Auth | Inherits `claude` CLI's credentials from env; nothing else needed |
| Tool naming | MCP tools surface to the agent as `mcp__<server-name>__<tool-name>` |
| Permissions | Tool ran without a permission round-trip (default policy permits MCP tools); `session/request_permission` handler in the test client was not exercised |

## Files

- `index.ts` — sidecar (~120 LOC). WSS server, per-connection child agent, ndjson relay, `session/new` interceptor.
- `test-client.ts` — headless ACP client (~170 LOC). Boots the sidecar, drives a session, asserts a `mcp__powerpoint__*` tool call.
- `package.json` — pinned to `@agentclientprotocol/claude-agent-acp@0.32.0`, `ws@^8.19.0`.

## Gotchas hit

1. **Node strip-only TS mode rejects parameter properties** (`constructor(private ws)`) — repo runs with `--experimental-strip-types`, so write fields explicitly.
2. **`auto-sideload` rewrites the PowerPoint manifest** when `--bridge` is on. Running the dev server with `BRIDGE_PORT=18080` re-sideloaded the add-in pointing at `:18080`. Restored to `0.5.0:8443` post-spike via `node scripts/sideload.mjs --tls`. Future automation should detect/skip auto-sideload when running side-by-side with another bridge instance.
3. **Two `:8080` listeners coexisted** (Claude Desktop's bridge + an unrelated `agentsvie` on IPv4) — not an issue, but explains an earlier surprise. Future taskpane integration should bind to a fixed alternate port.
4. **Cost** — 49k cached input tokens per session. Agent's stock system prompt + Claude Code tools list is heavy. The taskpane UI should reuse a single session across user turns rather than spinning a new one per prompt.

## What this clears

The Milestone-1 critical assumption from the design doc:

> If `claude-agent-acp` + `powerpoint-bridge`'s HTTP MCP actually compose end-to-end.

Cleared. Everything downstream (taskpane chat UI, Electron tray packaging, installer) is plumbing.

## Recommended next step (Milestone 2 — bare taskpane)

A second Office Add-in manifest co-resident with the existing bridge add-in. React/Vite scaffold. Connects to the sidecar over `wss://127.0.0.1:PORT/?token=…` (cert via mkcert). Renders only `agent_message_chunk` deltas + `tool_call` cards. Settings drawer wired to `_meta.systemPrompt`, `mcpServers`, allowed-tools list — all per-session injection points the sidecar already supports.

## Reverting the spike

```
git checkout main
git branch -D acp-spike       # if not merging
rm -rf sidecar/node_modules    # ~100 packages
```

The sidecar package is fully isolated from root `package.json` — nothing in the production install path was changed.
