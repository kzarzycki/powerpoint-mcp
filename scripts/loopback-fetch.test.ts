import { createServer } from 'node:http'
import { setTimeout as delay } from 'node:timers/promises'
import { expect, test } from 'vitest'
import { fetchLoopbackJson } from '../e2e/helpers/loopback-fetch.ts'

test('rejects when the server drops a partially received JSON response', async () => {
  // Real sockets need the response to arrive before disconnect; fake timers
  // cannot drive Node's network I/O. The race bounds a hung request.
  const server = createServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json', 'Content-Length': '100' })
    res.write('{')
    setTimeout(() => res.destroy(), 10)
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  try {
    const address = server.address()
    if (!address || typeof address === 'string') throw new Error('No TCP address')
    const result = fetchLoopbackJson(`http://127.0.0.1:${address.port}`, 5000).then(
      () => 'resolved',
      () => 'rejected',
    )
    expect(await Promise.race([result, delay(500, 'pending')])).toBe('rejected')
  } finally {
    await new Promise<void>((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())))
  }
})

test('bounds total request time even when partial data keeps arriving', async () => {
  // Exercise Node's real socket timeout versus AbortSignal's total deadline;
  // fake timers do not advance network I/O or the native abort timer.
  const server = createServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' })
    res.write('{')
    const timer = setInterval(() => res.write(' '), 10)
    res.on('close', () => clearInterval(timer))
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  try {
    const address = server.address()
    if (!address || typeof address === 'string') throw new Error('No TCP address')
    const result = fetchLoopbackJson(`http://127.0.0.1:${address.port}`, 100).then(
      () => 'resolved',
      () => 'rejected',
    )
    expect(await Promise.race([result, delay(500, 'pending')])).toBe('rejected')
  } finally {
    server.closeAllConnections()
    await new Promise<void>((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())))
  }
})
