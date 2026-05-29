#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
LOG_FILE="${POWERPOINT_MCP_LOG:-/tmp/powerpoint-mcp.log}"
MCP_PORT="${MCP_PORT:-3001}"
BRIDGE_PORT="${BRIDGE_PORT:-8080}"
TLS="${POWERPOINT_MCP_TLS:-0}"

if curl -sf "http://127.0.0.1:${MCP_PORT}/health" >/dev/null 2>&1 \
  && curl -sf "http://127.0.0.1:${BRIDGE_PORT}/health" >/dev/null 2>&1; then
  exit 0
fi

cd "$ROOT_DIR"
BRIDGE_TLS="$TLS" nohup node "$ROOT_DIR/dist/index.cjs" --http --bridge >>"$LOG_FILE" 2>&1 </dev/null &
disown || true
