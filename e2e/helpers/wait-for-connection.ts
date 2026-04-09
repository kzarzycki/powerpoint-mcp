import { ADDIN_CONNECT_TIMEOUT, E2E_BRIDGE_HEALTH, HEALTH_POLL_INTERVAL } from '../config.ts'

interface HealthResponse {
  status: string
  connections: number
}

/**
 * Poll the bridge /health endpoint until at least one add-in connection is active.
 * Throws with a diagnostic message if the timeout is reached.
 */
export async function waitForAddinConnection(timeoutMs = ADDIN_CONNECT_TIMEOUT): Promise<void> {
  const start = Date.now()
  let lastConnections = 0

  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(E2E_BRIDGE_HEALTH, { signal: AbortSignal.timeout(2000) })
      if (res.ok) {
        const body = (await res.json()) as HealthResponse
        lastConnections = body.connections
        if (lastConnections >= 1) return
      }
    } catch {
      // Bridge not responding yet
    }
    await new Promise((r) => setTimeout(r, HEALTH_POLL_INTERVAL))
  }

  throw new Error(
    `Add-in did not connect within ${timeoutMs}ms. ` +
      `Last known connections: ${lastConnections}. ` +
      'Possible causes: add-in taskpane not loading, WebSocket connection failing (cert trust issue?), ' +
      'or Office Web not processing the sideload URL parameters.',
  )
}
