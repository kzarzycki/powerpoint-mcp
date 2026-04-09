import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { test as base } from '@playwright/test'
import { E2E_MCP_URL } from '../config.ts'

export interface McpFixtures {
  mcpClient: Client
}

/**
 * Playwright fixture that provides a connected MCP client per test.
 * Connects via StreamableHTTP to the test bridge's MCP endpoint.
 */
export const test = base.extend<McpFixtures>({
  mcpClient: async ({}, use) => {
    const transport = new StreamableHTTPClientTransport(new URL(`${E2E_MCP_URL}/mcp`))
    const client = new Client({ name: 'e2e-test', version: '1.0.0' })
    await client.connect(transport)

    await use(client)

    await client.close()
  },
})

export { expect } from '@playwright/test'
