/**
 * Non-blocking startup version check against the npm registry.
 * Prints a one-line notice to stderr if a newer version is available.
 *
 * Disabled via BRIDGE_NO_UPDATE_CHECK=1 or --no-update-check flag.
 */

export interface VersionCheckOptions {
  currentVersion: string
  packageName?: string
  registryUrl?: string
  timeoutMs?: number
}

export interface VersionCheckResult {
  latest: string
  current: string
  updateAvailable: boolean
}

/**
 * Fetches the latest version from the npm registry and compares it to the
 * current version. Returns null on any error (network, parse, timeout).
 */
export async function checkForUpdate(options: VersionCheckOptions): Promise<VersionCheckResult | null> {
  const {
    currentVersion,
    packageName = 'powerpoint-mcp',
    registryUrl = 'https://registry.npmjs.org',
    timeoutMs = 3000,
  } = options

  try {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), timeoutMs)

    const res = await fetch(`${registryUrl}/${packageName}/latest`, {
      signal: controller.signal,
      headers: { Accept: 'application/json' },
    })
    clearTimeout(timer)

    if (!res.ok) return null

    const data = (await res.json()) as { version?: string }
    const latest = data.version
    if (!latest) return null

    return {
      latest,
      current: currentVersion,
      updateAvailable: isNewer(latest, currentVersion),
    }
  } catch {
    return null
  }
}

/**
 * Compares two semver strings. Returns true if `latest` is strictly newer than `current`.
 * Only compares major.minor.patch (ignores pre-release tags).
 */
export function isNewer(latest: string, current: string): boolean {
  const parse = (v: string) => v.replace(/^v/, '').split(/[-+]/)[0].split('.').map(Number)
  const [lMajor = 0, lMinor = 0, lPatch = 0] = parse(latest)
  const [cMajor = 0, cMinor = 0, cPatch = 0] = parse(current)

  if (lMajor !== cMajor) return lMajor > cMajor
  if (lMinor !== cMinor) return lMinor > cMinor
  return lPatch > cPatch
}

/**
 * Fire-and-forget version check. Logs to stderr if an update is available.
 * Never throws, never blocks startup.
 */
export function runVersionCheck(currentVersion: string): void {
  if (process.env.BRIDGE_NO_UPDATE_CHECK === '1') return
  if (process.argv.includes('--no-update-check')) return

  checkForUpdate({ currentVersion }).then((result) => {
    if (result?.updateAvailable) {
      console.error(
        `\n  Update available: v${result.current} → v${result.latest}` +
          '\n  Run "npm update -g powerpoint-mcp" to upgrade\n',
      )
    }
  })
}
