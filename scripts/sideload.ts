#!/usr/bin/env node

/**
 * Sideloads the add-in manifest into PowerPoint's WEF directory,
 * substituting the correct port from BRIDGE_PORT env var.
 *
 * Usage:
 *   node --experimental-strip-types scripts/sideload.ts        # HTTP manifest, default port 8080
 *   node --experimental-strip-types scripts/sideload.ts --tls  # HTTPS manifest, default port 8443
 *   BRIDGE_PORT=9090 node --experimental-strip-types scripts/sideload.ts  # Custom port
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { substituteManifestPort } from '../server/manifest.ts'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const PROJECT_ROOT = resolve(SCRIPT_DIR, '..')

const tls = process.argv.includes('--tls')
const defaultPort = tls ? 8443 : 8080
const port = Number(process.env.BRIDGE_PORT) || defaultPort

const manifestName = tls ? 'manifest-https.xml' : 'manifest.xml'
const src = resolve(PROJECT_ROOT, 'addin', manifestName)
const wefDir = resolve(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef')
const dest = resolve(wefDir, 'manifest.xml')

// Read template and substitute port (shared with the bridge server via substituteManifestPort)
const template = readFileSync(src, 'utf8')
const content = substituteManifestPort(template, defaultPort, port)

// Write to WEF directory
mkdirSync(wefDir, { recursive: true })
writeFileSync(dest, content)

// Write version marker
const pkg = JSON.parse(readFileSync(resolve(PROJECT_ROOT, 'package.json'), 'utf8'))
writeFileSync(resolve(PROJECT_ROOT, '.sideloaded'), `${pkg.version}:${port}`)

console.log(`[sideload] Manifest installed (${tls ? 'HTTPS' : 'HTTP'}, port ${port})`)
