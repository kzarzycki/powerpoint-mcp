import { randomUUID } from 'node:crypto'
import type { WebSocket } from 'ws'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface AddinConnection {
  ws: WebSocket
  ready: boolean
  presentationId: string
  filePath: string | null
}

interface PendingRequest {
  resolve: (data: unknown) => void
  reject: (err: Error) => void
  timer: ReturnType<typeof setTimeout>
  ws: WebSocket
}

// ---------------------------------------------------------------------------
// ConnectionPool — manages add-in WebSocket connections
// ---------------------------------------------------------------------------

export class ConnectionPool {
  private connections = new Map<string, AddinConnection>()
  private pendingRequests = new Map<string, PendingRequest>()
  private untitledCounter = 0
  private commandTimeout: number

  constructor(commandTimeout = 30_000) {
    this.commandTimeout = commandTimeout
  }

  get size(): number {
    return this.connections.size
  }

  add(presentationId: string, conn: AddinConnection): void {
    this.connections.set(presentationId, conn)
  }

  entries(): IterableIterator<[string, AddinConnection]> {
    return this.connections.entries()
  }

  /** Find which connection owns this WebSocket and remove it */
  removeBySocket(ws: WebSocket): string | null {
    for (const [id, conn] of this.connections) {
      if (conn.ws === ws) {
        this.connections.delete(id)
        return id
      }
    }
    return null
  }

  /** Reject all pending requests that were sent via a specific WebSocket */
  rejectPendingForSocket(ws: WebSocket): void {
    for (const [id, pending] of this.pendingRequests) {
      if (pending.ws === ws) {
        clearTimeout(pending.timer)
        pending.reject(new Error('Add-in disconnected'))
        this.pendingRequests.delete(id)
      }
    }
  }

  /** Handle an incoming response/error from the add-in */
  handleResponse(id: string, type: 'response' | 'error', data?: unknown, errorMessage?: string): void {
    const pending = this.pendingRequests.get(id)
    if (!pending) return
    clearTimeout(pending.timer)
    this.pendingRequests.delete(id)
    if (type === 'response') {
      pending.resolve(data)
    } else {
      pending.reject(new Error(errorMessage || 'Command failed'))
    }
  }

  /** Generate a presentation ID for a new connection */
  generateId(documentUrl: string | null): string {
    return documentUrl ?? `untitled-${++this.untitledCounter}`
  }

  /** Resolve which connection to target for a command */
  resolveTarget(presentationId?: string): AddinConnection {
    if (this.connections.size === 0) {
      throw new Error('No presentations connected. Open a PowerPoint file with the bridge add-in loaded.')
    }
    if (presentationId) {
      const conn = this.connections.get(presentationId)
      if (!conn)
        throw new Error(
          `Presentation not found: ${presentationId}. Use list_presentations to see connected presentations.`,
        )
      if (!conn.ready) throw new Error(`Presentation connected but not ready: ${presentationId}`)
      return conn
    }
    if (this.connections.size === 1) {
      const single = this.connections.values().next().value!
      if (!single.ready) throw new Error('Add-in connected but not ready')
      return single
    }
    const ids = [...this.connections.keys()]
    throw new Error(`Multiple presentations connected. Specify presentationId parameter. Available: ${ids.join(', ')}`)
  }

  /** Send a command to a specific WebSocket and wait for a response */
  sendCommand(
    action: string,
    params: Record<string, unknown>,
    targetWs: WebSocket,
    timeoutMs?: number,
  ): Promise<unknown> {
    const id = randomUUID()
    const timeout = timeoutMs ?? this.commandTimeout
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pendingRequests.delete(id)
        reject(new Error(`Command timed out after ${timeout}ms`))
      }, timeout)

      this.pendingRequests.set(id, { resolve, reject, timer, ws: targetWs })
      targetWs.send(JSON.stringify({ type: 'command', id, action, params }))
    })
  }
}
