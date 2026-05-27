# PowerPoint MCP

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D24.0.0-brightgreen)](https://nodejs.org)

An MCP server that lets AI assistants manipulate **live, open** PowerPoint presentations on macOS and PowerPoint Web via Office.js APIs.

Unlike file-based tools (python-pptx), PowerPoint MCP works with presentations that are already open — changes appear instantly, and you keep full access to PowerPoint's UI, animations, and formatting.

## Installation

### Claude Code Plugin (recommended)

Zero-config install from the marketplace. MCP auto-starts, add-in auto-sideloads, skill included.

```
/plugin marketplace add kzarzycki/powerpoint-mcp
/plugin install powerpoint-mcp@powerpoint-mcp
```

Then restart PowerPoint, open a presentation, and click the bridge add-in in the ribbon.

### Claude Desktop Extension

Build and install as a one-click `.mcpb` extension (from source):

```bash
git clone https://github.com/kzarzycki/powerpoint-mcp.git
cd powerpoint-mcp
npm install && npm run build:mcpb
open powerpoint-mcp-*.mcpb   # opens Claude Desktop installer
```

The extension auto-starts the bridge and auto-sideloads the add-in. Restart PowerPoint after first install.

**Known limitations:**
- **Chat mode only** — Cowork and Code tabs don't load desktop extensions ([upstream bug](https://github.com/anthropics/claude-code/issues/20377))
- **Single instance** — only one bridge can run on port 8080

### From source (development)

```bash
git clone https://github.com/kzarzycki/powerpoint-mcp.git
cd powerpoint-mcp
npm install
npm run sideload          # copies manifest to PowerPoint's add-in folder
npm start -- --bridge     # starts the add-in bridge (HTTP/WS on :8080) so the add-in can connect
```

Then restart PowerPoint, open a presentation, and click the bridge add-in in the ribbon.

## Motivation

This project was inspired by the [Claude in PowerPoint](https://support.anthropic.com/en/articles/11360939-using-claude-in-powerpoint) add-in. The first time I tried it, I was amazed — it edits live, open decks via Office.js, and the results are far better than file-based pptx tools. But it only works inside the add-in, which means no access to CLAUDE.md, skills, or any other Claude Code features. PowerPoint MCP brings those same Office.js capabilities to Claude Code (and any MCP client) so you get live editing with the full power of your coding environment.

> The project was originally named **powerpoint-bridge** and renamed to **powerpoint-mcp** in v0.5.0. The "bridge server" component keeps its name — it bridges MCP and the Office.js add-in over WebSocket.

## Architecture

```
AI Assistant  <--MCP STDIO/HTTP-->  Bridge Server (Node.js)  <--WS/WSS-->  PowerPoint Add-in (Office.js)
                                           |                                       |
                                     STDIO (default)                      Desktop: WKWebView sandbox
                                     or HTTP (:3001/mcp)                  Web: browser iframe
                                     localhost:8080 (HTTP)                Office.js API 1.1-1.10
                                     or :8443 (HTTPS)                    executes commands on
                                     serves add-in files + WS             live presentation
```

Two MCP transports are supported:
- **STDIO** (default) — used by plugin installs and `--stdio` flag; the MCP client spawns the server process directly
- **HTTP** — `127.0.0.1:3001/mcp`; enabled by the `--http` flag (e.g. `node server/index.ts --http --bridge` for development with `.mcp.json`)

Three components in one repo:

- **`addin/`** — Office.js taskpane add-in that loads inside PowerPoint and connects as a WebSocket client
- **`server/`** — Node.js bridge server: HTTP + WS + MCP transport (HTTPS/WSS opt-in via `BRIDGE_TLS=1`)
- **`skills/powerpoint-mcp/`** — Claude Code skill with tool docs, code patterns, and setup guide (auto-installed with plugin)
- **`certs/`** — Optional local TLS certificates for HTTPS mode (generated, gitignored)

## Prerequisites

- **macOS** (primary platform) or **PowerPoint Web** in Chrome/Brave (requires HTTPS mode)
- **Node.js >= 24** (uses native TypeScript execution)
- **Microsoft PowerPoint for Mac** or a **Microsoft 365** account for PowerPoint Web

```bash
brew install node
```

## Available Tools

| Tool | Description |
|------|-------------|
| `list_presentations` | Lists all connected presentations with their IDs and status |
| `inspect_deck` | Deck overview: slide dimensions, theme (colors + fonts), and all slides with layout and shape count |
| `inspect_layouts` | Slide layouts with names, types, placeholders, and positions. Supports field selection |
| `inspect_slide` | Returns detailed shape info for a slide (text, positions, sizes, fills) |
| `scan_slide` | Lightweight shape scanner: IDs, types, and positions |
| `screenshot_slide` | Captures a visual screenshot of a slide as PNG (requires PowerPoint 16.96+) |
| `preview_deck` | Returns thumbnails + text for all/selected slides in one call (efficient full-deck review) |
| `copy_slides` | Copies slides between two open presentations (data stays server-side, never in Claude context) |
| `insert_image` | Inserts an image from a file path, URL, or base64 data onto a slide |
| `get_local_copy` | Returns a local file path for the presentation (passthrough for local, exports cloud files to temp .pptx) |
| `search_fluent_icons` | Search Microsoft Fluent UI icons by keyword and insert as SVG |
| `execute_officejs` | Runs arbitrary Office.js code inside the live presentation |
| `add_slide` | Add a new slide from a named layout at a specific position, with placeholder text pre-filled |
| `duplicate_slide` | Duplicate a slide within the same presentation; data stays server-side |
| `read_deck_text` | Lightweight text extractor: slide titles and body text as plain strings (~20x smaller than inspect_slide) |
| `read_shape_paragraphs` | Read raw OOXML `<a:p>` paragraphs from a shape's text body, preserving all formatting |
| `edit_shape_paragraphs` | Replace paragraph content of a shape with raw OOXML `<a:p>` XML |
| `read_slide_xml` | Read the full raw OOXML of a slide, or filter to a specific shape |
| `edit_slide_xml` | Edit slide XML and reimport (xml mode or JS code mode for DOM manipulation) |
| `read_slide_zip` | Read multiple files from the exported slide zip (slide XML, rels, chart XMLs, Content_Types) |
| `edit_slide_zip` | Update multiple files in the slide zip and reimport in a single operation |
| `edit_slide_chart` | Create a chart on a slide from structured data (column, bar, line, pie, doughnut, area) via OOXML |
| `read_speaker_notes` | Read speaker notes from slides as plain text |
| `edit_speaker_notes` | Set speaker notes on one or more slides; accepts markdown text |
| `format_shapes` | Apply fill color and font formatting to multiple shapes on a slide in one call |
| `search_text` | Search for text across all slides — like grep for slides; supports regex |
| `verify_slides` | Run programmatic checks on a slide (overlaps, out-of-bounds, empty text, unused placeholders) |

When multiple presentations are open, pass `presentationId` (from `list_presentations`) to target a specific one.

## Limitations

- **Limited image control** — Images inserted via Common API (`insert_image` tool), not shape API; positioning works but no shape-level manipulation after insertion
- **Charts via OOXML** — Charts created by injecting OOXML (`edit_slide_chart`), not via Office.js chart API
- **No animations** — Not exposed in stable APIs
- **Solid fills only** — No gradients, effects, or shadows
- **Points for positioning** — All position/size values are in points (1 point = 1/72 inch)

## Security

PowerPoint MCP assumes a single-user, locally-trusted machine. Both network servers bind to the loopback interface only:

- The bridge server (HTTP + WebSocket) binds to `127.0.0.1:8080` (or `127.0.0.1:8443` with `BRIDGE_TLS=1`)
- The MCP HTTP transport binds to `127.0.0.1:3001`; the default STDIO transport opens no network port
- The MCP HTTP transport enables DNS-rebinding protection (rejects requests whose `Host` header is not `127.0.0.1:3001` / `localhost:3001`)
- No data leaves your machine

**Local-trust posture.** The add-in executes whatever JavaScript the bridge sends it over the local WebSocket, so any local process that can reach the loopback port can drive the open presentation. There is currently **no WebSocket authentication and no Origin allowlist** — loopback binding is the only access control. Do not run PowerPoint MCP on a shared or multi-user host.

**`execute_officejs` runs arbitrary code** inside PowerPoint's Office.js runtime. This is by design — it gives the AI full access to the Office.js API. Only use this with MCP clients you trust.

**Planned hardening.** Before binding to any non-loopback interface or shipping a remote-MCP deployment, the bridge needs a per-session handshake token, a WebSocket Origin allowlist, and an authenticated transport.

## Troubleshooting

**Add-in not appearing in PowerPoint**
1. Run `npm run sideload` and restart PowerPoint
2. Check that the file exists: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml`

**Add-in shows "Disconnected"**
Make sure the bridge server is running. In plugin mode the server auto-starts with Claude Code — call any tool to verify. For standalone installs, run `npm start -- --bridge` and verify with `curl http://127.0.0.1:8080/health`. The add-in auto-reconnects with exponential backoff.

**Using HTTPS mode (required for PowerPoint Web)**
HTTPS is required for PowerPoint Web and optional for desktop. To enable:
1. `brew install mkcert && mkcert -install` (one-time, trusts the CA in browsers)
2. `npm run setup-certs` to generate localhost certificates
3. For **desktop**: `npm run sideload:https` and restart PowerPoint
4. For **PowerPoint Web**: open a presentation at office.com, go to Home → Add-ins → Upload My Add-in, and upload `addin/manifest-https.xml`
5. Start the server: `BRIDGE_TLS=1 npm start -- --bridge`

## Platform Support

| Platform | Status |
|----------|--------|
| macOS | Supported (primary) |
| PowerPoint Web | Supported — requires HTTPS mode, sideload via browser (all tools work) |
| Windows | Untested — different sideloading path |
| Linux | Not supported (no PowerPoint for Linux) |

## Auto-Activation

The add-in uses `Office.AutoShowTaskpaneWithDocument` to minimize manual activation:

| Platform | Behavior |
|----------|----------|
| **macOS desktop** | Taskpane auto-opens when the add-in is sideloaded. Closing the pane keeps the WebSocket connection alive (shared runtime with `lifetime="long"`). |
| **PowerPoint Web** | One ribbon click per document per browser session to activate. After activation, the pane auto-reopens on page reload. Closing the pane keeps the connection alive within the session. |
| **New file from template** | OOXML template embedding can pre-activate the add-in (see below) |

Requires SharedRuntime 1.1 (PowerPoint 16.46+ on Mac, 2102+ on Windows). Centralized Deployment via Microsoft 365 admin center can eliminate the manual click for organizational use.

### Preparing Templates for Auto-Open

External tools can embed the add-in reference directly into `.pptx` templates so the add-in activates on first open without prior installation. This requires injecting two OOXML parts into the `.pptx` zip:

1. **`ppt/webextensions/webextension1.xml`** — add-in reference with `Office.AutoShowTaskpaneWithDocument` property
2. **`ppt/webextensions/taskpane.xml`** — taskpane configuration (dock state, visibility, width)

Plus the corresponding entries in `[Content_Types].xml` and relationship files.

Set `visibility="0"` if the add-in must already be sideloaded; set `visibility="1"` to prompt users to install it on first open.

For a complete working example, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin). The add-in ID to reference is `AE89909C-2813-4B08-9E1B-49E7379BD0E6` with `storeType="Registry"` and `store="developer"` for sideloaded installations.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup and guidelines.

## License

[MIT](LICENSE)
