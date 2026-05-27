import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Compact index entry: name + space-separated keywords */
interface RawIconEntry {
  /** Display name matching repo directory (e.g. "Warning", "Add Circle") */
  n: string
  /** Space-separated lowercase keywords for search */
  k: string
}

interface IconIndexEntry {
  name: string
  snakeName: string
  keywords: string[]
}

export interface IconSearchResult {
  id: string
  description: string
  isMono: boolean
  contentTier: string
  searchScore: number
  svgUrl: string
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const MANIFEST_URL =
  'https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/fonts/FluentSystemIcons-Regular.json'

const CDN_BASE = 'https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/assets'

const DEFAULT_SIZE = 24

// ---------------------------------------------------------------------------
// Caches
// ---------------------------------------------------------------------------

let cachedIndex: IconIndexEntry[] | null = null

// ---------------------------------------------------------------------------
// Name conversion helpers
// ---------------------------------------------------------------------------

function snakeToTitle(s: string): string {
  return s
    .split('_')
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ')
}

function titleToSnake(s: string): string {
  return s.replace(/\s+/g, '_').toLowerCase()
}

/** Convert display name to our icon ID: "Warning" → "Icons_Warning" */
export function nameToId(name: string, mono: boolean): string {
  const base = `Icons_${name.replace(/\s+/g, '_')}`
  return mono ? `${base}_M` : base
}

/** Parse icon ID back to snake name + mono flag */
export function parseIconId(iconId: string): { snakeName: string; isMono: boolean } {
  let id = iconId
  let isMono = false
  if (id.endsWith('_M')) {
    isMono = true
    id = id.slice(0, -2)
  }
  if (id.startsWith('Icons_')) {
    id = id.slice(6)
  }
  return { snakeName: id.toLowerCase(), isMono }
}

// ---------------------------------------------------------------------------
// Index loading
// ---------------------------------------------------------------------------

/** Build index from the font manifest JSON (keys are ic_fluent_{name}_{size}_regular) */
export function buildIndexFromManifest(manifest: Record<string, unknown>): IconIndexEntry[] {
  const seen = new Set<string>()
  const entries: IconIndexEntry[] = []

  for (const key of Object.keys(manifest)) {
    const match = key.match(/^ic_fluent_(.+)_(\d+)_regular$/)
    if (!match) continue
    const snakeName = match[1]
    if (seen.has(snakeName)) continue
    seen.add(snakeName)

    const name = snakeToTitle(snakeName)
    entries.push({
      name,
      snakeName,
      keywords: snakeName.split('_'),
    })
  }

  return entries.sort((a, b) => a.name.localeCompare(b.name))
}

/** Try loading static icon-index.json from disk */
function loadStaticIndex(): IconIndexEntry[] | null {
  try {
    const dir = dirname(fileURLToPath(import.meta.url))
    const raw = readFileSync(join(dir, 'icon-index.json'), 'utf-8')
    const data: RawIconEntry[] = JSON.parse(raw)
    return data.map((e) => ({
      name: e.n,
      snakeName: titleToSnake(e.n),
      keywords: e.k.split(' '),
    }))
  } catch {
    return null
  }
}

/** Load or fetch the icon index. Caches in memory after first call. */
export async function loadIndex(): Promise<IconIndexEntry[]> {
  if (cachedIndex) return cachedIndex

  // Try static JSON first
  const staticIndex = loadStaticIndex()
  if (staticIndex && staticIndex.length > 0) {
    cachedIndex = staticIndex
    return cachedIndex
  }

  // Fetch manifest from CDN
  const resp = await fetch(MANIFEST_URL)
  if (!resp.ok) {
    throw new Error(`Failed to fetch icon manifest: ${resp.status} ${resp.statusText}`)
  }
  const manifest = (await resp.json()) as Record<string, unknown>
  cachedIndex = buildIndexFromManifest(manifest)
  return cachedIndex
}

/** Reset cached index (for testing) */
export function resetIndex(): void {
  cachedIndex = null
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

function scoreMatch(entry: IconIndexEntry, queryWords: string[]): number {
  let score = 0
  const nameLower = entry.name.toLowerCase()
  const allWords = [nameLower, ...entry.keywords]

  for (const qw of queryWords) {
    // Exact keyword match
    if (entry.keywords.includes(qw)) {
      score += 10
      continue
    }
    // Name contains query word
    if (nameLower.includes(qw)) {
      score += 8
      continue
    }
    // Partial keyword match (keyword starts with query word)
    if (entry.keywords.some((kw) => kw.startsWith(qw))) {
      score += 5
      continue
    }
    // Any word contains query word as substring
    if (allWords.some((w) => w.includes(qw))) {
      score += 3
      continue
    }
    // No match for this query word — penalty
    score -= 2
  }

  // Bonus for exact name match
  const queryJoined = queryWords.join(' ')
  if (nameLower === queryJoined) {
    score += 20
  }

  return score
}

export async function searchIcons(query: string, top = 10, style?: 'regular' | 'filled'): Promise<IconSearchResult[]> {
  const index = await loadIndex()
  const queryWords = query
    .toLowerCase()
    .split(/\s+/)
    .filter((w) => w.length > 0)

  if (queryWords.length === 0) return []

  const scored = index
    .map((entry) => ({ entry, score: scoreMatch(entry, queryWords) }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, top)

  const isMono = style === 'regular'
  const isFilled = style === 'filled'

  // If no style filter, return both variants for each match. `scored` is
  // already capped at `top` icons above, so each icon expands to 2 rows here
  // (up to 2*top rows) — `top` means distinct icons, matching the style path.
  if (!isMono && !isFilled) {
    const results: IconSearchResult[] = []
    for (const { entry, score } of scored) {
      results.push({
        id: nameToId(entry.name, true),
        description: `${entry.name} (mono/outline)`,
        isMono: true,
        contentTier: 'free',
        searchScore: score,
        svgUrl: buildSvgUrl(entry.snakeName, true),
      })
      results.push({
        id: nameToId(entry.name, false),
        description: `${entry.name} (filled)`,
        isMono: false,
        contentTier: 'free',
        searchScore: score,
        svgUrl: buildSvgUrl(entry.snakeName, false),
      })
    }
    return results
  }

  return scored.map(({ entry, score }) => ({
    id: nameToId(entry.name, isMono),
    description: `${entry.name} (${isMono ? 'mono/outline' : 'filled'})`,
    isMono,
    contentTier: 'free',
    searchScore: score,
    svgUrl: buildSvgUrl(entry.snakeName, isMono),
  }))
}

// ---------------------------------------------------------------------------
// SVG fetch + recolor
// ---------------------------------------------------------------------------

export function buildSvgUrl(snakeName: string, isMono: boolean): string {
  const dirName = snakeToTitle(snakeName)
  const style = isMono ? 'regular' : 'filled'
  const fileName = `ic_fluent_${snakeName}_${DEFAULT_SIZE}_${style}.svg`
  return `${CDN_BASE}/${encodeURIComponent(dirName)}/SVG/${fileName}`
}

export function recolorSvg(svg: string, color: string): string {
  // Inject a CSS style block after the opening <svg> tag to override all fills
  const styleTag = `<style>.icon-color{fill:${color}}</style>`
  let result = svg.replace(/(<svg[^>]*>)/, `$1${styleTag}`)
  const SHAPE = '(?:path|circle|rect|polygon|ellipse)'
  // Element that already has a class AND a fill: merge icon-color into the
  // existing class value and drop the fill, so we never emit two class attrs.
  // Handle both attribute orders (class before fill, and fill before class).
  result = result.replace(
    new RegExp(`(<${SHAPE}[^>]*?)class="([^"]*)"([^>]*?)\\s*fill="[^"]*"`, 'g'),
    '$1class="$2 icon-color"$3',
  )
  result = result.replace(
    new RegExp(`(<${SHAPE}[^>]*?)fill="[^"]*"([^>]*?)\\s*class="([^"]*)"`, 'g'),
    '$1$2class="$3 icon-color"',
  )
  // Classless element with a fill: replace the fill with the class reference.
  result = result.replace(new RegExp(`(<${SHAPE}(?![^>]*class=)[^>]*?)fill="[^"]*"`, 'g'), '$1class="icon-color"')
  // For elements without a fill attribute, add the class
  result = result.replace(
    /(<(?:path|circle|rect|polygon|ellipse)(?![^>]*class=)[^>]*?)(\/?>)/g,
    '$1 class="icon-color"$2',
  )
  return result
}
