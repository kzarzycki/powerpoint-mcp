/** E2E test configuration constants */

export const E2E_BRIDGE_PORT = 9443
export const E2E_MCP_PORT = 9001
export const E2E_BRIDGE_URL = `https://localhost:${E2E_BRIDGE_PORT}`
export const E2E_MCP_URL = `http://localhost:${E2E_MCP_PORT}`
export const E2E_BRIDGE_HEALTH = `${E2E_BRIDGE_URL}/health`
export const E2E_MCP_HEALTH = `${E2E_MCP_URL}/health`

export const ADDIN_GUID = 'AE89909C-2813-4B08-9E1B-49E7379BD0E6'
export const ADDIN_MANIFEST_FILE = 'manifest-https.xml'

/** Build the sideload URL by appending Office dev query params to a document URL */
export function buildSideloadUrl(docUrl: string): string {
  const sep = docUrl.includes('?') ? '&' : '?'
  return [
    docUrl,
    `${sep}wdaddindevserverport=${E2E_BRIDGE_PORT}`,
    `&wdaddinmanifestfile=${ADDIN_MANIFEST_FILE}`,
    `&wdaddinmanifestguid=${ADDIN_GUID}`,
    '&wdaddintest=true',
  ].join('')
}

/** Browser profile directory for persistent login sessions */
export const BROWSER_PROFILE_DIR = new URL('.browser-profile', import.meta.url).pathname

/** Timeouts */
export const SERVER_START_TIMEOUT = 15_000
export const ADDIN_CONNECT_TIMEOUT = 60_000
export const HEALTH_POLL_INTERVAL = 500
