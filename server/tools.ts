import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import type { ConnectionPool } from './bridge.ts'
import { registerChartTools } from './tools/charts.ts'
import { registerExecTools } from './tools/exec.ts'
import { registerInspectTools } from './tools/inspect.ts'
import { registerMediaTools } from './tools/media.ts'
import { registerNotesTools } from './tools/notes.ts'
import { registerSlideTools } from './tools/slides.ts'
import { registerTextTools } from './tools/text.ts'
import { registerXmlTools } from './tools/xml.ts'

export type { FormatShapeSpec } from './tools/shared.ts'
// Re-exports for external importers (server/index.ts, tools.test.ts) and the
// grouped tool modules. The actual definitions live in ./tools/shared.ts.
export {
  buildFormatShapeOps,
  buildInsertOptions,
  clearSessionWarnings,
  getConcurrentWarning,
  globToRegExp,
  localCopyCache,
  parseSlideRange,
  themeCache,
} from './tools/shared.ts'

// ---------------------------------------------------------------------------
// Tool registration — thin aggregator over the grouped modules in ./tools/.
// ---------------------------------------------------------------------------

export function registerTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  registerInspectTools(server, pool, getSessionId, getActiveSessionCount)
  registerSlideTools(server, pool, getSessionId, getActiveSessionCount)
  registerTextTools(server, pool, getSessionId, getActiveSessionCount)
  registerXmlTools(server, pool, getSessionId, getActiveSessionCount)
  registerNotesTools(server, pool, getSessionId, getActiveSessionCount)
  registerChartTools(server, pool, getSessionId, getActiveSessionCount)
  registerMediaTools(server, pool, getSessionId, getActiveSessionCount)
  registerExecTools(server, pool, getSessionId, getActiveSessionCount)
}
