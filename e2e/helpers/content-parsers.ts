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

/** Parse the text content of an MCP tool result as JSON */
export function getJsonContent<T = unknown>(result: ToolResult): T {
  const text = getTextContent(result)
  return JSON.parse(text) as T
}

/** Check if the tool result indicates an error */
export function isToolError(result: ToolResult): boolean {
  return 'isError' in result && result.isError === true
}
