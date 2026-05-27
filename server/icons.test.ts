import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildIndexFromManifest, nameToId, parseIconId, resetIndex, searchIcons } from './icons.ts'

// Mock fs so loadStaticIndex() fails and we control the index via buildIndexFromManifest
vi.mock('node:fs', () => ({
  readFileSync: vi.fn(() => {
    throw new Error('no static index')
  }),
  existsSync: vi.fn(() => false),
  writeFileSync: vi.fn(),
}))

// Sample manifest matching fluentui-system-icons format
const SAMPLE_MANIFEST: Record<string, number> = {
  ic_fluent_warning_16_regular: 0xe001,
  ic_fluent_warning_20_regular: 0xe002,
  ic_fluent_warning_24_regular: 0xe003,
  ic_fluent_arrow_down_20_regular: 0xe004,
  ic_fluent_arrow_down_24_regular: 0xe005,
  ic_fluent_checkmark_circle_20_regular: 0xe006,
  ic_fluent_checkmark_circle_24_regular: 0xe007,
  ic_fluent_lightbulb_20_regular: 0xe008,
  ic_fluent_lightbulb_24_regular: 0xe009,
  ic_fluent_shield_20_regular: 0xe00a,
  ic_fluent_shield_24_regular: 0xe00b,
  ic_fluent_rocket_20_regular: 0xe00c,
  ic_fluent_rocket_24_regular: 0xe00d,
  ic_fluent_heart_20_regular: 0xe00e,
  ic_fluent_heart_24_regular: 0xe00f,
  ic_fluent_star_20_regular: 0xe010,
  ic_fluent_star_24_regular: 0xe011,
  ic_fluent_home_20_regular: 0xe012,
  ic_fluent_home_24_regular: 0xe013,
  ic_fluent_settings_20_regular: 0xe014,
  ic_fluent_settings_24_regular: 0xe015,
}

// Mock global fetch for both manifest and SVG fetching
const mockFetch = vi.fn()
vi.stubGlobal('fetch', mockFetch)

describe('icons', () => {
  afterEach(() => {
    resetIndex()
    mockFetch.mockReset()
  })

  describe('buildIndexFromManifest', () => {
    it('extracts unique icon names from font manifest', () => {
      const index = buildIndexFromManifest(SAMPLE_MANIFEST)
      const names = index.map((e) => e.name)
      expect(names).toEqual([
        'Arrow Down',
        'Checkmark Circle',
        'Heart',
        'Home',
        'Lightbulb',
        'Rocket',
        'Settings',
        'Shield',
        'Star',
        'Warning',
      ])
    })

    it('deduplicates entries with different sizes', () => {
      const index = buildIndexFromManifest(SAMPLE_MANIFEST)
      // Warning has 3 size variants but should appear once
      const warningEntries = index.filter((e) => e.name === 'Warning')
      expect(warningEntries).toHaveLength(1)
    })

    it('generates correct snake names and keywords', () => {
      const index = buildIndexFromManifest(SAMPLE_MANIFEST)
      const checkmark = index.find((e) => e.name === 'Checkmark Circle')
      expect(checkmark?.snakeName).toBe('checkmark_circle')
      expect(checkmark?.keywords).toEqual(['checkmark', 'circle'])
    })

    it('ignores non-regular entries', () => {
      const manifest = {
        ...SAMPLE_MANIFEST,
        ic_fluent_alert_24_filled: 0xf001, // filled, not regular — should be ignored
        some_other_key: 0xf002, // doesn't match pattern
      }
      const index = buildIndexFromManifest(manifest)
      const names = index.map((e) => e.name)
      expect(names).not.toContain('Alert')
    })
  })

  describe('nameToId', () => {
    it('creates filled icon ID', () => {
      expect(nameToId('Warning', false)).toBe('Icons_Warning')
    })

    it('creates mono icon ID with _M suffix', () => {
      expect(nameToId('Warning', true)).toBe('Icons_Warning_M')
    })

    it('handles multi-word names', () => {
      expect(nameToId('Arrow Down', false)).toBe('Icons_Arrow_Down')
      expect(nameToId('Arrow Down', true)).toBe('Icons_Arrow_Down_M')
    })
  })

  describe('parseIconId', () => {
    it('parses filled icon ID', () => {
      expect(parseIconId('Icons_Warning')).toEqual({
        snakeName: 'warning',
        isMono: false,
      })
    })

    it('parses mono icon ID', () => {
      expect(parseIconId('Icons_Warning_M')).toEqual({
        snakeName: 'warning',
        isMono: true,
      })
    })

    it('handles multi-word icon IDs', () => {
      expect(parseIconId('Icons_Arrow_Down')).toEqual({
        snakeName: 'arrow_down',
        isMono: false,
      })
      expect(parseIconId('Icons_Arrow_Down_M')).toEqual({
        snakeName: 'arrow_down',
        isMono: true,
      })
    })
  })

  describe('searchIcons', () => {
    function setupManifestFetch() {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => SAMPLE_MANIFEST,
      })
    }

    it('returns matching icons sorted by score', async () => {
      setupManifestFetch()
      const results = await searchIcons('warning', 5)
      expect(results.length).toBeGreaterThan(0)
      expect(results[0].id).toContain('Warning')
    })

    it('returns both mono and filled variants by default', async () => {
      setupManifestFetch()
      const results = await searchIcons('warning', 10)
      const monoResult = results.find((r) => r.isMono)
      const filledResult = results.find((r) => !r.isMono)
      expect(monoResult).toBeDefined()
      expect(filledResult).toBeDefined()
    })

    it('filters to regular style only', async () => {
      setupManifestFetch()
      const results = await searchIcons('warning', 10, 'regular')
      for (const r of results) {
        expect(r.isMono).toBe(true)
      }
    })

    it('filters to filled style only', async () => {
      setupManifestFetch()
      const results = await searchIcons('warning', 10, 'filled')
      for (const r of results) {
        expect(r.isMono).toBe(false)
      }
    })

    it('respects top limit', async () => {
      setupManifestFetch()
      const results = await searchIcons('a', 3)
      expect(results.length).toBeLessThanOrEqual(3)
    })

    it('returns empty for no match', async () => {
      setupManifestFetch()
      const results = await searchIcons('xyznonexistent', 5)
      expect(results).toEqual([])
    })

    it('returns empty for empty query', async () => {
      setupManifestFetch()
      const results = await searchIcons('', 5)
      expect(results).toEqual([])
    })

    it('includes contentTier, searchScore, and svgUrl', async () => {
      setupManifestFetch()
      const results = await searchIcons('heart', 2)
      expect(results[0].contentTier).toBe('free')
      expect(results[0].searchScore).toBeGreaterThan(0)
      expect(results[0].svgUrl).toMatch(/^https:\/\/raw\.githubusercontent\.com\/microsoft\/fluentui-system-icons/)
    })

    it('caches index after first load', async () => {
      setupManifestFetch()
      await searchIcons('warning', 1)
      // Second call should not trigger another fetch
      const results = await searchIcons('star', 1)
      expect(results.length).toBeGreaterThan(0)
      expect(mockFetch).toHaveBeenCalledTimes(1)
    })
  })
})
