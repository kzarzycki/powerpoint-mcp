import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { get as httpGet } from 'node:http'
import { get as httpsGet } from 'node:https'
import { join } from 'node:path'

// The e2e bridge's HTTPS cert (certs/localhost.pem) is a leaf certificate
// signed by mkcert's local development CA, not self-signed. Node's fetch/https
// don't consult the system keychain trust store `mkcert -install` populates,
// so the parent process needs that CA explicitly to validate the chain. This
// extends trust to one already-locally-trusted CA; it never disables
// certificate verification.
let cachedCa: Buffer | undefined

function mkcertRootCa(): Buffer {
  if (cachedCa) return cachedCa
  const caRoot = execFileSync('mkcert', ['-CAROOT'], { encoding: 'utf8' }).trim()
  cachedCa = readFileSync(join(caRoot, 'rootCA.pem'))
  return cachedCa
}

/**
 * Fetch JSON from a loopback URL the e2e bridge is serving. For https, trusts
 * mkcert's local CA (see above) so the mkcert-signed dev cert validates
 * normally instead of being rejected as unknown.
 */
export function fetchLoopbackJson<T>(url: string, timeoutMs = 5000): Promise<T> {
  const { promise, resolve, reject } = Promise.withResolvers<T>()
  const get = url.startsWith('https:') ? httpsGet : httpGet
  const req = get(url, url.startsWith('https:') ? { ca: mkcertRootCa() } : {}, (res) => {
    const chunks: Buffer[] = []
    res.on('data', (chunk: Buffer) => chunks.push(chunk))
    res.on('end', () => {
      try {
        if (res.statusCode !== 200) throw new Error(`status ${res.statusCode}`)
        resolve(JSON.parse(Buffer.concat(chunks).toString()) as T)
      } catch (error) {
        reject(error instanceof Error ? error : new Error(String(error)))
      }
    })
  })
  req.on('error', reject)
  req.setTimeout(timeoutMs, () => req.destroy(new Error('timed out')))
  return promise
}
