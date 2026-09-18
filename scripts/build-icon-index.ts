#!/usr/bin/env node
/**
 * Fetches icon names from microsoft/fluentui-system-icons GitHub repo
 * and generates server/icon-index.json for search_icons/insert_icon tools.
 *
 * Usage: npx tsx scripts/build-icon-index.ts
 */

import { writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const OUTPUT_PATH = join(__dirname, '..', 'server', 'icon-index.json')

const GITHUB_API = 'https://api.github.com'
const REPO = 'microsoft/fluentui-system-icons'

interface GitTreeEntry {
  path: string
  type: 'tree' | 'blob'
  sha: string
}

interface IconEntry {
  /** Display name, matches repo directory (e.g. "Warning", "Add Circle") */
  n: string
  /** Space-separated lowercase keywords for search */
  k: string
}

function generateKeywords(name: string): string {
  // Split PascalCase and spaces into individual words
  const words = name
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .split(/[\s_]+/)
    .map((w) => w.toLowerCase())
    .filter((w) => w.length > 0)
  return [...new Set(words)].join(' ')
}

async function fetchJson(url: string): Promise<unknown> {
  const headers: Record<string, string> = {
    Accept: 'application/vnd.github+json',
    'User-Agent': 'powerpoint-mcp-icon-index-builder',
  }
  if (process.env.GITHUB_TOKEN) {
    headers.Authorization = `Bearer ${process.env.GITHUB_TOKEN}`
  }
  const resp = await fetch(url, { headers })
  if (!resp.ok) {
    throw new Error(`GitHub API error: ${resp.status} ${resp.statusText} for ${url}`)
  }
  return resp.json()
}

async function main() {
  console.log('Fetching repo tree...')

  // Get root tree to find assets directory SHA
  const rootTree = (await fetchJson(`${GITHUB_API}/repos/${REPO}/git/trees/main`)) as {
    tree: GitTreeEntry[]
  }
  const assetsEntry = rootTree.tree.find((e) => e.path === 'assets' && e.type === 'tree')
  if (!assetsEntry) {
    throw new Error('Could not find assets directory in repo tree')
  }

  // Get all icon directories from assets tree
  console.log('Fetching assets tree...')
  const assetsTree = (await fetchJson(`${GITHUB_API}/repos/${REPO}/git/trees/${assetsEntry.sha}`)) as {
    tree: GitTreeEntry[]
    truncated: boolean
  }

  if (assetsTree.truncated) {
    console.warn('Warning: assets tree was truncated, some icons may be missing')
  }

  const icons: IconEntry[] = assetsTree.tree
    .filter((e) => e.type === 'tree')
    .map((e) => ({
      n: e.path,
      k: generateKeywords(e.path),
    }))
    .sort((a, b) => a.n.localeCompare(b.n))

  console.log(`Found ${icons.length} icons`)

  writeFileSync(OUTPUT_PATH, JSON.stringify(icons))
  console.log(`Written to ${OUTPUT_PATH}`)
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})
