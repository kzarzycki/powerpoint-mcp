import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

interface Harness {
  context: Record<string, unknown>
  sockets: FakeSocket[]
  timers: Map<number, () => void>
  delays: number[]
  ready: (info?: Record<string, string>) => void
  runTimer: () => void
}

class FakeSocket {
  onopen: (() => void) | null = null
  onclose: (() => void) | null = null
  onerror: ((error: unknown) => void) | null = null
  onmessage: ((event: { data: string }) => void) | null = null
  sent: string[] = []

  send(message: string): void {
    this.sent.push(message)
  }
  open(): void {
    this.onopen?.()
  }
  close(): void {
    this.onclose?.()
  }
}

const JITTER = 0.25

function harness(): Harness {
  const sockets: FakeSocket[] = []
  const timers = new Map<number, () => void>()
  const delays: number[] = []
  let nextTimer = 1
  let ready: ((info?: Record<string, string>) => void) | undefined
  const document = { getElementById: () => ({ textContent: '', className: '' }) }
  const office = {
    onReady: (callback: (info?: Record<string, string>) => void) => {
      ready = callback
    },
    context: {
      document: { url: 'file:///disposable.pptx', settings: { set: () => undefined, saveAsync: () => undefined } },
      requirements: { isSetSupported: () => false },
    },
    addin: { setStartupBehavior: () => undefined },
    StartupBehavior: { load: 'load' },
  }
  const context: Record<string, unknown> = {
    Office: office,
    WebSocket: class extends FakeSocket {
      constructor() {
        super()
        sockets.push(this)
      }
    },
    window: { location: { protocol: 'http:', host: 'localhost:8080', origin: 'http://localhost:8080' } },
    document,
    console: { log: () => undefined, error: () => undefined },
    Math: Object.assign(Object.create(Math), { random: () => JITTER }),
    setTimeout: (callback: () => void, delay: number) => {
      const id = nextTimer++
      timers.set(id, callback)
      delays.push(delay)
      return id
    },
    clearTimeout: (id: number) => {
      timers.delete(id)
    },
    PowerPoint: undefined,
  }
  const source = readFileSync(join(process.cwd(), 'addin/app.js'), 'utf8')
  runInNewContext(source, context)
  return {
    context,
    sockets,
    timers,
    delays,
    ready: (info = { host: 'PowerPoint', platform: 'mac' }) => ready?.(info),
    runTimer: () => {
      const [id, callback] = timers.entries().next().value ?? []
      if (id !== undefined) timers.delete(id)
      callback?.()
    },
  }
}

describe('add-in WebSocket ownership', () => {
  it.each([
    [
      'Office ready before fallback',
      (h: Harness) => {
        h.ready()
        h.runTimer()
      },
    ],
    [
      'Office ready after fallback',
      (h: Harness) => {
        h.runTimer()
        h.ready()
      },
    ],
    [
      'plain browser without ready callback',
      (h: Harness) => {
        h.runTimer()
      },
    ],
  ])('%s opens one socket', (_name, setup) => {
    const h = harness()
    setup(h)
    expect(h.sockets).toHaveLength(1)
  })

  it('does not add a socket while one is connecting or open', () => {
    const h = harness()
    h.runTimer()
    const connect = h.context.connect as () => void
    connect()
    h.sockets[0].open()
    connect()
    expect(h.sockets).toHaveLength(1)
  })

  it('uses exponential backoff plus jitter, caps before jitter, and resets after open', () => {
    const h = harness()
    h.runTimer()
    const jitter = Math.floor(JITTER * 1000)
    const expected = [500, 1000, 2000, 4000, 8000, 16000, 30000].map((delay) => delay + jitter)
    for (const delay of expected) {
      h.sockets.at(-1)?.close()
      expect(h.delays.at(-1)).toBe(delay)
      h.runTimer()
    }
    h.sockets.at(-1)?.open()
    h.sockets.at(-1)?.close()
    expect(h.delays.at(-1)).toBe(500 + jitter)
    h.runTimer()
    expect(h.sockets).toHaveLength(expected.length + 2)
  })

  it('cancels the pending reconnect timer when connect runs first', () => {
    const h = harness()
    h.runTimer()
    h.sockets[0].close()
    expect(h.timers.size).toBe(1)
    const connect = h.context.connect as () => void
    connect()
    expect(h.timers.size).toBe(0)
    expect(h.sockets).toHaveLength(2)
  })

  it('ignores a captured stale timer after a replacement socket exists', () => {
    const h = harness()
    h.runTimer()
    h.sockets[0].close()
    const stale = [...h.timers.values()].at(-1)
    const connect = h.context.connect as () => void
    connect()
    stale?.()
    expect(h.sockets).toHaveLength(2)
  })
})
