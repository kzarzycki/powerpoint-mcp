import type { Client } from '@modelcontextprotocol/sdk/client/index.js'

/** The actual return type of client.callTool() */
type ToolResult = Awaited<ReturnType<Client['callTool']>>

/** Extract all text content from an MCP tool result, joined by newlines */
export function getTextContent(result: ToolResult): string {
  if (!('content' in result)) return ''
  return (result.content as Array<{ type: string; text?: string }>)
    .filter((c) => c.type === 'text' && c.text)
    .map((c) => c.text!)
    .join('\n')
}

/**
 * Parse the text content of an MCP tool result as JSON.
 *
 * Some tools append a human-readable note after the JSON payload (e.g. the
 * concurrent-session warning from inspect_deck). Parse the leading JSON object/array
 * and ignore any trailing text.
 */
export function getJsonContent<T = unknown>(result: ToolResult): T {
  const text = getTextContent(result)
  try {
    return JSON.parse(text) as T
  } catch {
    const end = Math.max(text.lastIndexOf('}'), text.lastIndexOf(']'))
    if (end !== -1) return JSON.parse(text.slice(0, end + 1)) as T
    throw new Error(`Tool result is not JSON: ${text.slice(0, 200)}`)
  }
}

/** Check if the tool result indicates an error */
export function isToolError(result: ToolResult): boolean {
  return 'isError' in result && result.isError === true
}
