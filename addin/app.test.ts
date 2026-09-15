import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

interface Harness {
  context: Record<string, unknown>
  sockets: FakeSocket[]
  timers: Array<() => void>
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

function harness(): Harness {
  const sockets: FakeSocket[] = []
  const timers: Array<() => void> = []
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
    setTimeout: (callback: () => void) => {
      timers.push(callback)
      return timers.length - 1
    },
    clearTimeout: (id: number) => {
      timers[id] = () => undefined
    },
    PowerPoint: undefined,
  }
  const source = readFileSync(join(process.cwd(), 'addin/app.js'), 'utf8')
  runInNewContext(source, context)
  return {
    context,
    sockets,
    timers,
    ready: (info = { host: 'PowerPoint', platform: 'mac' }) => ready?.(info),
    runTimer: () => timers.shift()?.(),
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

  it('reconnects a closed socket and ignores its stale timer', () => {
    const h = harness()
    h.runTimer()
    h.sockets[0].close()
    const connect = h.context.connect as () => void
    connect()
    expect(h.sockets).toHaveLength(2)
    h.runTimer()
    expect(h.sockets).toHaveLength(2)
  })
})
