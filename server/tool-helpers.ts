import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'

/**
 * What a tool handler may return:
 *  - a `string`  → passed through unchanged as a single text content item
 *    (handlers build their own text, including any concurrent-access warning suffix)
 *  - a pre-built `CallToolResult` envelope (e.g. image + text) → passed through as-is
 *  - any other value (object/array/number/etc.) → pretty-printed as JSON text
 */
type ToolReturn = string | CallToolResult | object

/** Extract a human-readable message from an unknown thrown value. */
export function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

function isCallToolResult(value: unknown): value is CallToolResult {
  return typeof value === 'object' && value !== null && Array.isArray((value as { content?: unknown }).content)
}

/**
 * Wraps a tool handler with the standard error→result conversion and
 * success formatting, removing the duplicated try/catch epilogue from every
 * handler. The wrapped function keeps the handler's `(args, extra)` signature
 * and always resolves to a `CallToolResult`.
 */
export function withTool<Args, Extra>(
  handler: (args: Args, extra: Extra) => ToolReturn | Promise<ToolReturn>,
): (args?: Args, extra?: Extra) => Promise<CallToolResult> {
  return async (args?: Args, extra?: Extra) => {
    try {
      const result = await handler(args as Args, extra as Extra)
      if (typeof result === 'string') {
        return { content: [{ type: 'text', text: result }] }
      }
      if (isCallToolResult(result)) {
        return result
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] }
    } catch (err: unknown) {
      return { content: [{ type: 'text', text: `Error: ${errorMessage(err)}` }], isError: true }
    }
  }
}
