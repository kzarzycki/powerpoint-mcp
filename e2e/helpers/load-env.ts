import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const CONFIG_FILE = resolve(import.meta.dirname, '..', 'local-config.json')
const ENV_FILE = resolve(import.meta.dirname, '..', '.env')

export function loadE2eEnv(): void {
  if (existsSync(CONFIG_FILE)) {
    const config = JSON.parse(readFileSync(CONFIG_FILE, 'utf8')) as Record<string, string>
    for (const [key, value] of Object.entries(config)) {
      if (process.env[key] === undefined) process.env[key] = value
    }
    return
  }
  if (!existsSync(ENV_FILE)) return
  const content = readFileSync(ENV_FILE, 'utf8')
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) continue
    const match = trimmed.match(/^([A-Z_][A-Z0-9_]*)=(.*)$/)
    if (!match) continue
    const [, key, rawValue] = match
    if (process.env[key] !== undefined) continue
    process.env[key] = rawValue.replace(/^["']|["']$/g, '')
  }
}
