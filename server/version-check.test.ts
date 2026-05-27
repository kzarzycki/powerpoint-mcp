import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { checkForUpdate, isNewer, runVersionCheck, type VersionCheckOptions } from './version-check.ts'

describe('isNewer', () => {
  it('returns true when latest has higher major', () => {
    expect(isNewer('2.0.0', '1.0.0')).toBe(true)
  })

  it('returns true when latest has higher minor', () => {
    expect(isNewer('1.2.0', '1.1.0')).toBe(true)
  })

  it('returns true when latest has higher patch', () => {
    expect(isNewer('1.0.2', '1.0.1')).toBe(true)
  })

  it('returns false when versions are equal', () => {
    expect(isNewer('1.0.0', '1.0.0')).toBe(false)
  })

  it('returns false when current is newer', () => {
    expect(isNewer('1.0.0', '1.0.1')).toBe(false)
  })

  it('handles v prefix', () => {
    expect(isNewer('v2.0.0', 'v1.0.0')).toBe(true)
  })

  it('ignores pre-release tags', () => {
    expect(isNewer('1.0.1-beta.1', '1.0.0')).toBe(true)
  })

  it('ignores build metadata when comparing', () => {
    expect(isNewer('1.2.4+build', '1.2.3')).toBe(true)
    expect(isNewer('1.2.3+build', '1.2.3')).toBe(false)
  })
})

describe('checkForUpdate', () => {
  const baseOpts: VersionCheckOptions = {
    currentVersion: '0.4.0',
    packageName: 'powerpoint-mcp',
    timeoutMs: 1000,
  }

  beforeEach(() => {
    vi.stubGlobal('fetch', vi.fn())
  })

  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('returns updateAvailable=true when registry has newer version', async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: true,
      json: async () => ({ version: '1.0.0' }),
    } as Response)

    const result = await checkForUpdate(baseOpts)
    expect(result).toEqual({
      latest: '1.0.0',
      current: '0.4.0',
      updateAvailable: true,
    })
  })

  it('returns updateAvailable=false when versions match', async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: true,
      json: async () => ({ version: '0.4.0' }),
    } as Response)

    const result = await checkForUpdate(baseOpts)
    expect(result).toEqual({
      latest: '0.4.0',
      current: '0.4.0',
      updateAvailable: false,
    })
  })

  it('returns null on network error', async () => {
    vi.mocked(fetch).mockRejectedValue(new Error('ENOTFOUND'))

    const result = await checkForUpdate(baseOpts)
    expect(result).toBeNull()
  })

  it('returns null on non-ok response', async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: false,
      status: 404,
    } as Response)

    const result = await checkForUpdate(baseOpts)
    expect(result).toBeNull()
  })

  it('returns null when response has no version field', async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: true,
      json: async () => ({}),
    } as Response)

    const result = await checkForUpdate(baseOpts)
    expect(result).toBeNull()
  })

  it('uses custom registry URL', async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: true,
      json: async () => ({ version: '0.4.0' }),
    } as Response)

    await checkForUpdate({ ...baseOpts, registryUrl: 'https://custom.registry.com' })
    expect(fetch).toHaveBeenCalledWith('https://custom.registry.com/powerpoint-mcp/latest', expect.any(Object))
  })
})

describe('runVersionCheck', () => {
  const originalEnv = process.env.BRIDGE_NO_UPDATE_CHECK

  beforeEach(() => {
    vi.stubGlobal('fetch', vi.fn())
    delete process.env.BRIDGE_NO_UPDATE_CHECK
  })

  afterEach(() => {
    vi.restoreAllMocks()
    if (originalEnv !== undefined) {
      process.env.BRIDGE_NO_UPDATE_CHECK = originalEnv
    } else {
      delete process.env.BRIDGE_NO_UPDATE_CHECK
    }
  })

  it('skips check when BRIDGE_NO_UPDATE_CHECK=1', () => {
    process.env.BRIDGE_NO_UPDATE_CHECK = '1'
    runVersionCheck('0.4.0')
    expect(fetch).not.toHaveBeenCalled()
  })

  it('fires fetch when env var is not set', () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: true,
      json: async () => ({ version: '0.4.0' }),
    } as Response)

    runVersionCheck('0.4.0')
    expect(fetch).toHaveBeenCalled()
  })
})
