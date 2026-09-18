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

/**
 * Whether a command whose outcome could not be observed directly (timeout or
 * disconnect) never reached the add-in at all ("never-started" — safe to
 * retry) or was sent and its result is unknown ("unknown" — must never be
 * retried automatically, since it may have already mutated the deck).
 */
export type CommandOutcome = 'never-started' | 'unknown'

/** Rejection raised when a command's definitive result could not be observed. */
export class CommandOutcomeError extends Error {
  readonly outcome: CommandOutcome

  constructor(outcome: CommandOutcome, message: string) {
    super(message)
    this.name = 'CommandOutcomeError'
    this.outcome = outcome
  }
}

/** Evidence that a response arrived after its caller already gave up — never includes presentation content. */
export interface LateCompletion {
  id: string
  type: 'response' | 'error'
  at: number
}

export interface ConnectionLiveness {
  lastSeenAt: number | null
  alive: boolean
}

interface QueuedCommand {
  id: string
  action: string
  params: Record<string, unknown>
  ws: WebSocket
  timeoutMs: number
  resolve: (data: unknown) => void
  reject: (err: Error) => void
}

interface DispatchedCommand {
  id: string
  ws: WebSocket
  timeoutMs: number
  resolve: (data: unknown) => void
  reject: (err: Error) => void
  timer: ReturnType<typeof setTimeout>
}

/** One execution slot per socket: at most one command in flight, the rest wait their turn. */
interface SocketQueue {
  current: DispatchedCommand | null
  waiting: QueuedCommand[]
}

interface GraceTombstone {
  ws: WebSocket
  timer: ReturnType<typeof setTimeout>
}

const LATE_COMPLETION_LOG_LIMIT = 50

// ---------------------------------------------------------------------------
// ConnectionPool — manages add-in WebSocket connections
// ---------------------------------------------------------------------------

export class ConnectionPool {
  private connections = new Map<string, AddinConnection>()
  private queues = new Map<WebSocket, SocketQueue>()
  private pendingRequests = new Map<string, DispatchedCommand>()
  private graceTombstones = new Map<string, GraceTombstone>()
  private lateCompletions: LateCompletion[] = []
  private lastSeenAt = new Map<WebSocket, number>()
  private commandTimeout: number

  constructor(commandTimeout = 30_000) {
    this.commandTimeout = commandTimeout
  }

  get size(): number {
    return this.connections.size
  }

  /**
   * Register a connection under `presentationId`. Safe to repeat: the same
   * socket registering again under the same id is a no-op replace. A socket
   * registering under a *new* id (e.g. after a document identity changes)
   * has its stale entry removed first, so a reconnect or re-registration
   * never leaves a phantom duplicate behind.
   */
  add(presentationId: string, conn: AddinConnection): void {
    for (const [id, existing] of this.connections) {
      if (existing.ws === conn.ws && id !== presentationId) {
        this.connections.delete(id)
      }
    }
    this.connections.set(presentationId, conn)
    this.markSeen(conn.ws)
  }

  entries(): IterableIterator<[string, AddinConnection]> {
    return this.connections.entries()
  }

  /** Find every connection this WebSocket owns and remove it; returns the first removed ID for logging. */
  removeBySocket(ws: WebSocket): string | null {
    let removedId: string | null = null
    for (const [id, conn] of this.connections) {
      if (conn.ws === ws) {
        this.connections.delete(id)
        removedId ??= id
      }
    }
    return removedId
  }

  /**
   * Tear down a socket's command queue: the in-flight command (if any) settles
   * as "unknown" (it may still be running in the add-in) and every command
   * still waiting behind it settles as "never-started" (it was never sent).
   * A late completion for the in-flight command remains recordable during its
   * grace window even though the socket is gone, since it may reconnect.
   */
  rejectPendingForSocket(ws: WebSocket): void {
    this.lastSeenAt.delete(ws)
    const queue = this.queues.get(ws)
    if (!queue) return

    if (queue.current) {
      this.settleAsUnknown(queue.current, 'Add-in disconnected')
    }
    for (const queued of queue.waiting) {
      queued.reject(new CommandOutcomeError('never-started', `Add-in disconnected before ${queued.action} started`))
    }
    queue.waiting = []
    this.queues.delete(ws)
  }

  /** Handle an incoming response/error from the add-in. */
  handleResponse(id: string, type: 'response' | 'error', data?: unknown, errorMessage?: string): void {
    const pending = this.pendingRequests.get(id)
    if (pending) {
      clearTimeout(pending.timer)
      this.pendingRequests.delete(id)
      this.markSeen(pending.ws)
      if (type === 'response') {
        pending.resolve(data)
      } else {
        pending.reject(new Error(errorMessage || 'Command failed'))
      }
      this.advanceQueue(pending.ws)
      return
    }

    const tomb = this.graceTombstones.get(id)
    if (tomb) {
      clearTimeout(tomb.timer)
      this.graceTombstones.delete(id)
      this.markSeen(tomb.ws)
      this.recordLateCompletion(id, type)
      this.advanceQueue(tomb.ws)
      return
    }

    // Unknown or already-recorded id: still evidence of a completion the
    // caller can no longer be attached to. Never silently dropped.
    this.recordLateCompletion(id, type)
  }

  /** Evidence log of completions that arrived after their caller had already been given a definitive outcome. Never includes response/error content. */
  getLateCompletions(): LateCompletion[] {
    return [...this.lateCompletions]
  }

  /**
   * Generate a presentation ID for a `ready` message. A saved file's URL is
   * itself a stable identity, so the same URL always yields the same ID
   * across a reconnect. An unsaved deck has no stable identity signal beyond
   * "no URL", so the smallest untitled slot not currently occupied is reused
   * — freed by a disconnect before the same runtime's next `ready` arrives —
   * while any still-open unsaved deck keeps its own slot.
   */
  generateId(documentUrl: string | null): string {
    if (documentUrl !== null) return documentUrl
    let n = 1
    while (this.connections.has(`untitled-${n}`)) n++
    return `untitled-${n}`
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

  /** Record that a socket produced traffic (registration or a command outcome) just now. */
  markSeen(ws: WebSocket): void {
    this.lastSeenAt.set(ws, Date.now())
  }

  /** Liveness evidence for a presentation: when it was last seen, and whether that is within `maxAgeMs`. */
  getLiveness(presentationId: string, maxAgeMs: number): ConnectionLiveness {
    const conn = this.connections.get(presentationId)
    if (!conn) throw new Error(`Presentation not found: ${presentationId}`)
    const lastSeenAt = this.lastSeenAt.get(conn.ws) ?? null
    const alive = lastSeenAt !== null && Date.now() - lastSeenAt <= maxAgeMs
    return { lastSeenAt, alive }
  }

  isAlive(presentationId: string, maxAgeMs: number): boolean {
    return this.getLiveness(presentationId, maxAgeMs).alive
  }

  /**
   * Send a command to a specific WebSocket and wait for a response. Commands
   * to the same socket execute one at a time: a command queued behind
   * another is only sent — and only starts its own execution timeout — once
   * the socket's current command has settled.
   */
  sendCommand(
    action: string,
    params: Record<string, unknown>,
    targetWs: WebSocket,
    timeoutMs?: number,
  ): Promise<unknown> {
    const id = randomUUID()
    const timeout = timeoutMs ?? this.commandTimeout
    const { promise, resolve, reject } = Promise.withResolvers<unknown>()
    const queued: QueuedCommand = { id, action, params, ws: targetWs, timeoutMs: timeout, resolve, reject }
    const queue = this.queues.get(targetWs) ?? { current: null, waiting: [] }
    this.queues.set(targetWs, queue)
    if (queue.current === null) {
      this.dispatch(queue, queued)
    } else {
      queue.waiting.push(queued)
    }
    return promise
  }

  // -- internals -------------------------------------------------------------

  private dispatch(queue: SocketQueue, queued: QueuedCommand): void {
    const dispatched: DispatchedCommand = {
      id: queued.id,
      ws: queued.ws,
      timeoutMs: queued.timeoutMs,
      resolve: queued.resolve,
      reject: queued.reject,
      timer: setTimeout(() => this.onExecutionTimeout(queued.id), queued.timeoutMs),
    }
    queue.current = dispatched
    this.pendingRequests.set(queued.id, dispatched)
    queued.ws.send(JSON.stringify({ type: 'command', id: queued.id, action: queued.action, params: queued.params }))
  }

  private onExecutionTimeout(id: string): void {
    const dispatched = this.pendingRequests.get(id)
    if (!dispatched) return
    this.pendingRequests.delete(id)
    this.settleAsUnknown(dispatched, `Command timed out after ${dispatched.timeoutMs}ms`)
  }

  /** Reject a dispatched command as outcome-unknown and open a grace window for a possible late completion. */
  private settleAsUnknown(dispatched: DispatchedCommand, message: string): void {
    clearTimeout(dispatched.timer)
    this.pendingRequests.delete(dispatched.id)
    dispatched.reject(new CommandOutcomeError('unknown', message))

    const timer = setTimeout(() => {
      this.graceTombstones.delete(dispatched.id)
      this.advanceQueue(dispatched.ws)
    }, dispatched.timeoutMs)
    this.graceTombstones.set(dispatched.id, { ws: dispatched.ws, timer })
  }

  /** Free a socket's execution slot and dispatch the next queued command, if any. */
  private advanceQueue(ws: WebSocket): void {
    const queue = this.queues.get(ws)
    if (!queue) return
    queue.current = null
    const next = queue.waiting.shift()
    if (next) this.dispatch(queue, next)
  }

  private recordLateCompletion(id: string, type: 'response' | 'error'): void {
    this.lateCompletions.push({ id, type, at: Date.now() })
    if (this.lateCompletions.length > LATE_COMPLETION_LOG_LIMIT) this.lateCompletions.shift()
  }
}
