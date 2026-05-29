#!/bin/bash
# This script is for standalone/development use only.
# Plugin users: run `claude plugin add kzarzycki/powerpoint-mcp` instead.
set -e

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"

echo "=== PowerPoint MCP Setup ==="
echo ""

# 1. Sideload add-in manifest (HTTP by default, honours BRIDGE_PORT)
node --experimental-strip-types "$REPO_DIR/scripts/sideload.ts"

# 2. Install skill globally (skip if running as a Claude Code plugin)
if [ -z "${CLAUDE_PLUGIN_ROOT:-}" ]; then
  mkdir -p ~/.claude/skills
  ln -sfn "$REPO_DIR/skills/powerpoint-mcp" ~/.claude/skills/powerpoint-mcp
  echo "[skill] Installed globally at ~/.claude/skills/powerpoint-mcp"
else
  echo "[skill] Skipped (plugin auto-discovery handles this)"
fi

echo ""
echo "=== Setup complete ==="
echo ""
echo "Next steps:"
echo "  1. Restart PowerPoint to load the add-in"
echo "  2. Start the bridge: npm start -- --bridge"
echo "  3. In any project, ask Claude: 'enable powerpoint mcp in this project'"
echo ""
echo "Optional: To use HTTPS/WSS instead of HTTP/WS:"
echo "  brew install mkcert && mkcert -install"
echo "  npm run setup-certs"
echo "  npm run sideload:https"
echo "  BRIDGE_TLS=1 npm start -- --bridge"
