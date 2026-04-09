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

### npx (any MCP client)

Run as an MCP server via npx (no install needed):

```bash
npx powerpoint-mcp --stdio --bridge
```

Then configure your MCP client:

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "npx",
      "args": ["-y", "powerpoint-mcp", "--stdio", "--bridge"]
    }
  }
}
```

**Claude Code** (`.mcp.json` in project root):
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "powerpoint-mcp", "--stdio", "--bridge"]
    }
  }
}
```

**Cursor** (`.cursor/mcp.json`):
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "npx",
      "args": ["-y", "powerpoint-mcp", "--stdio", "--bridge"]
    }
  }
}
```

**VS Code / GitHub Copilot** (`.vscode/mcp.json`):
```json
{
  "servers": {
    "powerpoint-mcp": {
      "command": "npx",
      "args": ["-y", "powerpoint-mcp", "--stdio", "--bridge"]
    }
  }
}
```

**Windsurf** (`~/.codeium/windsurf/mcp_config.json`):
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "command": "npx",
      "args": ["-y", "powerpoint-mcp", "--stdio", "--bridge"]
    }
  }
}
```

> **Note:** The PowerPoint add-in must still be sideloaded separately. Run `npx powerpoint-mcp --sideload` or see [Troubleshooting](#troubleshooting).

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
npm run sideload     # copies manifest to PowerPoint's add-in folder
npm start            # starts MCP server (STDIO mode by default)
```

Then restart PowerPoint, open a presentation, and click the bridge add-in in the ribbon.

## Motivation

This project was inspired by the [Claude in PowerPoint](https://support.anthropic.com/en/articles/11360939-using-claude-in-powerpoint) add-in. The first time I tried it, I was amazed — it edits live, open decks via Office.js, and the results are far better than file-based pptx tools. But it only works inside the add-in, which means no access to CLAUDE.md, skills, or any other Claude Code features. PowerPoint MCP brings those same Office.js capabilities to Claude Code (and any MCP client) so you get live editing with the full power of your coding environment.

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
- **HTTP** — `localhost:3001/mcp`; used by `npm start` for standalone/development setups

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

When multiple presentations are open, pass `presentationId` (from `list_presentations`) to target a specific one.

## Limitations

- **Limited image control** — Images inserted via Common API (`insert_image` tool), not shape API; positioning works but no shape-level manipulation after insertion
- **Charts via OOXML** — Charts created by injecting OOXML (`edit_slide_chart`), not via Office.js chart API
- **No animations** — Not exposed in stable APIs
- **Solid fills only** — No gradients, effects, or shadows
- **Points for positioning** — All position/size values are in points (1 point = 1/72 inch)

## Security

PowerPoint MCP runs entirely on localhost:

- The bridge server binds to `localhost:8080` (HTTP) or `localhost:8443` (HTTPS with `BRIDGE_TLS=1`)
- MCP transport is either STDIO (no network port) or HTTP on `localhost:3001`
- No data leaves your machine

**`execute_officejs` runs arbitrary code** inside PowerPoint's Office.js runtime. This is by design — it gives the AI full access to the Office.js API. Only use this with MCP clients you trust.

## Troubleshooting

**Add-in not appearing in PowerPoint**
1. Run `npm run sideload` (or `npx powerpoint-mcp --sideload`) and restart PowerPoint
2. Check that the file exists: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml`

**Add-in shows "Disconnected"**
Make sure the bridge server is running. In plugin mode the server auto-starts with Claude Code — call any tool to verify. For standalone installs, run `npm start` and verify with `curl http://localhost:8080/health`. The add-in auto-reconnects with exponential backoff.

**Using HTTPS mode (required for PowerPoint Web)**
HTTPS is required for PowerPoint Web and optional for desktop. To enable:
1. `brew install mkcert && mkcert -install` (one-time, trusts the CA in browsers)
2. `npm run setup-certs` to generate localhost certificates
3. For **desktop**: `npm run sideload:https` and restart PowerPoint
4. For **PowerPoint Web**: open a presentation at office.com, go to Home → Add-ins → Upload My Add-in, and upload `addin/manifest-https.xml`
5. Start the server: `BRIDGE_TLS=1 npm start`

## Platform Support

| Platform | Status |
|----------|--------|
| macOS | Supported (primary) |
| PowerPoint Web | Supported — requires HTTPS mode, sideload via browser (all 23 tools work) |
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
