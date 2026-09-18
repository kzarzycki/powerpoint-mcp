import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import type { AddinConnection } from './bridge.ts'
import { ConnectionPool } from './bridge.ts'

function mockWs(overrides: Partial<WebSocket> = {}): WebSocket {
  return {
    send: vi.fn(),
    readyState: 1,
    ...overrides,
  } as unknown as WebSocket
}

function makeConn(ws: WebSocket, opts: Partial<AddinConnection> = {}): AddinConnection {
  return {
    ws,
    ready: true,
    presentationId: opts.presentationId ?? 'test.pptx',
    filePath: opts.filePath ?? '/path/test.pptx',
    ...opts,
  }
}

interface SentCommand {
  id: string
  type: string
  action: string
  params: Record<string, unknown>
}

/** Reads the Nth JSON command a mock socket's `send` was called with, without an unchecked cast. */
function sentCommand(ws: WebSocket, index = 0): SentCommand {
  const [payload] = vi.mocked(ws.send).mock.calls[index]
  if (typeof payload !== 'string') throw new Error('expected a string ws.send payload')
  return JSON.parse(payload)
}

describe('ConnectionPool', () => {
  let pool: ConnectionPool

  beforeEach(() => {
    pool = new ConnectionPool(100) // 100ms timeout for fast tests
  })

  describe('resolveTarget', () => {
    it('throws when no connections exist', () => {
      expect(() => pool.resolveTarget()).toThrow('No presentations connected')
    })

    it('returns the single connection when only one exists', () => {
      const ws = mockWs()
      const conn = makeConn(ws)
      pool.add('test.pptx', conn)
      expect(pool.resolveTarget()).toBe(conn)
    })

    it('throws when single connection is not ready', () => {
      const ws = mockWs()
      pool.add('test.pptx', makeConn(ws, { ready: false }))
      expect(() => pool.resolveTarget()).toThrow('not ready')
    })

    it('throws listing available IDs when multiple connections exist without ID', () => {
      pool.add('a.pptx', makeConn(mockWs(), { presentationId: 'a.pptx' }))
      pool.add('b.pptx', makeConn(mockWs(), { presentationId: 'b.pptx' }))
      expect(() => pool.resolveTarget()).toThrow('Multiple presentations connected')
      expect(() => pool.resolveTarget()).toThrow('a.pptx')
      expect(() => pool.resolveTarget()).toThrow('b.pptx')
    })

    it('returns correct connection when presentationId is specified', () => {
      const wsA = mockWs()
      const wsB = mockWs()
      const connA = makeConn(wsA, { presentationId: 'a.pptx' })
      const connB = makeConn(wsB, { presentationId: 'b.pptx' })
      pool.add('a.pptx', connA)
      pool.add('b.pptx', connB)
      expect(pool.resolveTarget('b.pptx')).toBe(connB)
    })

    it('throws when specified presentationId is not found', () => {
      pool.add('a.pptx', makeConn(mockWs()))
      expect(() => pool.resolveTarget('missing.pptx')).toThrow('Presentation not found: missing.pptx')
    })

    it('throws when specified presentation is not ready', () => {
      pool.add('a.pptx', makeConn(mockWs(), { ready: false, presentationId: 'a.pptx' }))
      expect(() => pool.resolveTarget('a.pptx')).toThrow('not ready')
    })
  })

  describe('removeBySocket', () => {
    it('removes the connection and returns its ID', () => {
      const ws = mockWs()
      pool.add('test.pptx', makeConn(ws))
      expect(pool.size).toBe(1)
      const removed = pool.removeBySocket(ws)
      expect(removed).toBe('test.pptx')
      expect(pool.size).toBe(0)
    })

    it('returns null when socket is not found', () => {
      expect(pool.removeBySocket(mockWs())).toBeNull()
    })
  })

  describe('generateId', () => {
    it('returns documentUrl when provided', () => {
      expect(pool.generateId('/path/to/file.pptx')).toBe('/path/to/file.pptx')
    })

    it('assigns distinct ids to concurrently open unsaved decks', () => {
      const first = pool.generateId(null)
      pool.add(first, makeConn(mockWs(), { presentationId: first, filePath: null }))
      const second = pool.generateId(null)
      expect(second).not.toBe(first)
    })

    it('reuses a freed untitled slot once its connection disconnects', () => {
      const ws = mockWs()
      const first = pool.generateId(null)
      pool.add(first, makeConn(ws, { presentationId: first, filePath: null }))
      pool.removeBySocket(ws)

      const reused = pool.generateId(null)
      expect(reused).toBe(first)
    })
  })

  describe('sendCommand', () => {
    it('sends JSON command and resolves when response arrives', async () => {
      const ws = mockWs()
      const promise = pool.sendCommand('executeCode', { code: 'test' }, ws)

      // Extract the command ID from what was sent
      const sentJson = sentCommand(ws)
      expect(sentJson.type).toBe('command')
      expect(sentJson.action).toBe('executeCode')

      // Simulate response
      pool.handleResponse(sentJson.id, 'response', { result: 42 })
      await expect(promise).resolves.toEqual({ result: 42 })
    })

    it('rejects when error response arrives', async () => {
      const ws = mockWs()
      const promise = pool.sendCommand('executeCode', { code: 'bad' }, ws)

      const sentJson = sentCommand(ws)
      pool.handleResponse(sentJson.id, 'error', undefined, 'Something went wrong')

      await expect(promise).rejects.toThrow('Something went wrong')
    })

    it('rejects on timeout', async () => {
      vi.useFakeTimers()
      try {
        const ws = mockWs()
        const promise = pool.sendCommand('executeCode', { code: 'slow' }, ws)

        vi.advanceTimersByTime(200) // past the 100ms timeout
        await expect(promise).rejects.toThrow('Command timed out')
      } finally {
        vi.useRealTimers()
      }
    })
  })

  describe('rejectPendingForSocket', () => {
    it('rejects only pending requests for the given socket', async () => {
      const wsA = mockWs()
      const wsB = mockWs()

      const promiseA = pool.sendCommand('executeCode', { code: 'a' }, wsA)
      const promiseB = pool.sendCommand('executeCode', { code: 'b' }, wsB)

      pool.rejectPendingForSocket(wsA)

      await expect(promiseA).rejects.toThrow('Add-in disconnected')

      // promiseB should still be pending — resolve it manually
      const sentB = sentCommand(wsB)
      pool.handleResponse(sentB.id, 'response', 'ok')
      await expect(promiseB).resolves.toBe('ok')
    })

    it('marks the in-flight command unknown and every still-queued command never-started', async () => {
      const ws = mockWs()
      const running = pool.sendCommand('executeCode', { code: 'running' }, ws)
      const queued = pool.sendCommand('executeCode', { code: 'queued' }, ws)
      expect(ws.send).toHaveBeenCalledTimes(1)

      pool.rejectPendingForSocket(ws)

      await expect(running).rejects.toMatchObject({ outcome: 'unknown' })
      await expect(queued).rejects.toMatchObject({ outcome: 'never-started' })
      expect(ws.send).toHaveBeenCalledTimes(1)
    })
  })

  describe('command lifecycle', () => {
    it('starts the execution timeout only when a queued command is actually sent', async () => {
      vi.useFakeTimers()
      try {
        const ws = mockWs()
        const first = pool.sendCommand('executeCode', { code: 'first' }, ws)
        const second = pool.sendCommand('executeCode', { code: 'second' }, ws)
        expect(ws.send).toHaveBeenCalledTimes(1)

        vi.advanceTimersByTime(99)
        await Promise.resolve()
        expect(ws.send).toHaveBeenCalledTimes(1)

        const firstJson = sentCommand(ws, 0)
        pool.handleResponse(firstJson.id, 'response', 'first-result')
        await expect(first).resolves.toBe('first-result')
        expect(ws.send).toHaveBeenCalledTimes(2)

        vi.advanceTimersByTime(99)
        await Promise.resolve()
        expect(ws.send).toHaveBeenCalledTimes(2)

        vi.advanceTimersByTime(1)
        await expect(second).rejects.toMatchObject({ outcome: 'unknown' })
      } finally {
        vi.useRealTimers()
      }
    })

    it('records a late completion without retaining its data and never resolves twice', async () => {
      vi.useFakeTimers()
      try {
        const ws = mockWs()
        const promise = pool.sendCommand('executeCode', { code: 'mutate once' }, ws)
        const sent = sentCommand(ws)

        vi.advanceTimersByTime(100)
        await expect(promise).rejects.toMatchObject({ outcome: 'unknown' })
        expect(ws.send).toHaveBeenCalledTimes(1)

        pool.handleResponse(sent.id, 'response', { secretPresentationContent: 'do not retain' })
        const completions = pool.getLateCompletions()
        expect(completions).toHaveLength(1)
        expect(completions[0]).toMatchObject({ id: sent.id, type: 'response' })
        expect(JSON.stringify(completions)).not.toContain('do not retain')

        // A second, duplicate late response for the same id must not throw or double-record.
        pool.handleResponse(sent.id, 'response', { more: 'data' })
        expect(pool.getLateCompletions()).toHaveLength(2)
        expect(ws.send).toHaveBeenCalledTimes(1) // never retried
      } finally {
        vi.useRealTimers()
      }
    })

    it('releases a lost completion only after its grace period, then continues the queue', async () => {
      vi.useFakeTimers()
      try {
        const ws = mockWs()
        const first = pool.sendCommand('executeCode', { code: 'lost' }, ws)
        const second = pool.sendCommand('executeCode', { code: 'next' }, ws)
        vi.advanceTimersByTime(100)
        await expect(first).rejects.toMatchObject({ outcome: 'unknown' })
        expect(ws.send).toHaveBeenCalledTimes(1)

        vi.advanceTimersByTime(99)
        expect(ws.send).toHaveBeenCalledTimes(1)
        vi.advanceTimersByTime(1)
        expect(ws.send).toHaveBeenCalledTimes(2)

        const secondJson = sentCommand(ws, 1)
        pool.handleResponse(secondJson.id, 'response', 'next-result')
        await expect(second).resolves.toBe('next-result')
      } finally {
        vi.useRealTimers()
      }
    })

    it('frees the execution slot immediately on a definitive response, not after a grace period', async () => {
      vi.useFakeTimers()
      try {
        const ws = mockWs()
        const first = pool.sendCommand('executeCode', { code: 'first' }, ws)
        const second = pool.sendCommand('executeCode', { code: 'second' }, ws)

        const firstJson = sentCommand(ws, 0)
        pool.handleResponse(firstJson.id, 'response', 'ok')
        await expect(first).resolves.toBe('ok')

        // The next command dispatches right away — no grace-period wait for a definitive outcome.
        expect(ws.send).toHaveBeenCalledTimes(2)
        void second
      } finally {
        vi.useRealTimers()
      }
    })
  })

  describe('registration identity', () => {
    it('keeps one entry and one identifier when the same document reconnects on a new socket', () => {
      const url = '/path/first.pptx'
      const firstWs = mockWs()
      const firstId = pool.generateId(url)
      pool.add(firstId, makeConn(firstWs, { presentationId: firstId, filePath: url }))

      const newWs = mockWs()
      const reconnectId = pool.generateId(url)
      expect(reconnectId).toBe(firstId)
      const replacement = makeConn(newWs, { presentationId: reconnectId, filePath: url })
      pool.add(reconnectId, replacement)

      expect(pool.size).toBe(1)
      expect(pool.resolveTarget(firstId)).toBe(replacement)
    })

    it('registering the same socket under a new identity removes the stale entry', () => {
      const ws = mockWs()
      pool.add('a.pptx', makeConn(ws, { presentationId: 'a.pptx', filePath: '/path/a.pptx' }))
      pool.add('b.pptx', makeConn(ws, { presentationId: 'b.pptx', filePath: '/path/b.pptx' }))

      expect(pool.size).toBe(1)
      expect([...pool.entries()].map(([id]) => id)).toEqual(['b.pptx'])
      expect(pool.resolveTarget('b.pptx').filePath).toBe('/path/b.pptx')
    })

    it('is safe to repeat registration for the same socket and id', () => {
      const ws = mockWs()
      const conn = makeConn(ws, { presentationId: 'a.pptx', filePath: '/path/a.pptx' })
      pool.add('a.pptx', conn)
      pool.add('a.pptx', conn)
      expect(pool.size).toBe(1)
      expect(pool.resolveTarget('a.pptx')).toBe(conn)
    })

    it('keeps genuinely different presentations separately addressable', () => {
      const firstId = pool.generateId('/path/first.pptx')
      const secondId = pool.generateId('/path/second.pptx')
      pool.add(firstId, makeConn(mockWs(), { presentationId: firstId, filePath: '/path/first.pptx' }))
      pool.add(secondId, makeConn(mockWs(), { presentationId: secondId, filePath: '/path/second.pptx' }))

      expect(pool.size).toBe(2)
      expect(pool.resolveTarget(firstId).filePath).toBe('/path/first.pptx')
      expect(pool.resolveTarget(secondId).filePath).toBe('/path/second.pptx')
    })
  })

  describe('liveness evidence', () => {
    it('tracks last-seen activity and reports a stale connection as not alive', () => {
      vi.useFakeTimers()
      try {
        vi.setSystemTime(new Date('2026-01-01T00:00:00.000Z'))
        const ws = mockWs()
        pool.add('deck', makeConn(ws, { presentationId: 'deck' }))
        const initial = pool.getLiveness('deck', 100)
        expect(initial.lastSeenAt).toBe(Date.now())
        expect(initial.alive).toBe(true)

        vi.advanceTimersByTime(50)
        pool.markSeen(ws)
        expect(pool.getLiveness('deck', 100).lastSeenAt).toBe(Date.now())

        vi.advanceTimersByTime(101)
        expect(pool.isAlive('deck', 100)).toBe(false)
      } finally {
        vi.useRealTimers()
      }
    })

    it('marks a connection seen when its command response arrives', async () => {
      vi.useFakeTimers()
      try {
        vi.setSystemTime(new Date('2026-01-01T00:00:00.000Z'))
        const ws = mockWs()
        pool.add('deck', makeConn(ws, { presentationId: 'deck' }))
        vi.advanceTimersByTime(10)
        const promise = pool.sendCommand('executeCode', { code: 'read' }, ws)
        const sent = sentCommand(ws)
        pool.handleResponse(sent.id, 'response', 'ok')
        await expect(promise).resolves.toBe('ok')
        expect(pool.getLiveness('deck', 100).lastSeenAt).toBe(Date.now())
      } finally {
        vi.useRealTimers()
      }
    })
  })
})
