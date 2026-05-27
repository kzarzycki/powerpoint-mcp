import { existsSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ConnectionPool } from '../bridge.ts'
import type { ThemeInfo } from '../xml-helpers.ts'

// ---------------------------------------------------------------------------
// Shared state, pure helpers, and cross-group Office.js helpers used by the
// grouped tool modules. registerTools wires the groups together in tools.ts.
// ---------------------------------------------------------------------------

export const localCopyCache = new Map<string, { localPath: string; revision: number }>()
export const themeCache = new Map<string, ThemeInfo>()

// ---------------------------------------------------------------------------
// Concurrent access warning helper
// ---------------------------------------------------------------------------

const sessionConcurrentWarnings = new Map<string, Set<string>>()

export function getConcurrentWarning(
  mcpSessionId: string | undefined,
  presentationId: string,
  activeSessions: number,
): string | null {
  if (!mcpSessionId) return null
  if (activeSessions <= 1) return null

  const warned = sessionConcurrentWarnings.get(mcpSessionId)
  if (warned?.has(presentationId)) return null

  if (!warned) {
    sessionConcurrentWarnings.set(mcpSessionId, new Set([presentationId]))
  } else {
    warned.add(presentationId)
  }

  return '\n\nNote: Other MCP sessions are also connected to the bridge. If they target this presentation, changes apply immediately (last-write-wins).'
}

export function clearSessionWarnings(sessionId: string): void {
  sessionConcurrentWarnings.delete(sessionId)
}

export function parseSlideRange(range: string | undefined): number[] | null {
  if (!range) return null
  const indices = new Set<number>()
  for (const part of range.split(',')) {
    const trimmed = part.trim()
    if (!trimmed) continue
    const dashIdx = trimmed.indexOf('-', 1)
    if (dashIdx === -1) {
      const n = Number(trimmed)
      if (!Number.isInteger(n) || n < 0) throw new Error(`Invalid slide index: "${trimmed}"`)
      indices.add(n)
    } else {
      const start = Number(trimmed.slice(0, dashIdx))
      const end = Number(trimmed.slice(dashIdx + 1))
      if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start) {
        throw new Error(`Invalid slide range: "${trimmed}"`)
      }
      for (let i = start; i <= end; i++) indices.add(i)
    }
  }
  if (indices.size === 0) return null
  return [...indices].sort((a, b) => a - b)
}

export interface FormatShapeSpec {
  id: string
  fill?: string
  font?: {
    bold?: boolean
    italic?: boolean
    size?: number
    color?: string
    name?: string
  }
}

/**
 * Build the per-shape Office.js op string for format_shapes. User-supplied
 * strings (id, fill, font.color, font.name) are JSON.stringify'd so embedded
 * quotes can't break out of the generated string literals.
 */
export function buildFormatShapeOps(shapes: FormatShapeSpec[], slideIndex: number): string {
  return shapes
    .map((s) => {
      const lines: string[] = []
      lines.push(`  var s = shapeMap[${JSON.stringify(s.id)}];`)
      lines.push(`  if (!s) throw new Error("Shape " + ${JSON.stringify(s.id)} + " not found on slide ${slideIndex}");`)

      if (s.fill) {
        lines.push(`  s.fill.setSolidColor(${JSON.stringify(s.fill)});`)
      }

      if (s.font) {
        lines.push(`  var tf = s.getTextFrameOrNullObject();`)
        lines.push(`  await context.sync();`)
        lines.push(`  if (!tf.isNullObject) {`)
        lines.push(`    var tr = tf.textRange;`)
        if (s.font.bold !== undefined) lines.push(`    tr.font.bold = ${s.font.bold};`)
        if (s.font.italic !== undefined) lines.push(`    tr.font.italic = ${s.font.italic};`)
        if (s.font.size !== undefined) lines.push(`    tr.font.size = ${s.font.size};`)
        if (s.font.color !== undefined) lines.push(`    tr.font.color = ${JSON.stringify(s.font.color)};`)
        if (s.font.name !== undefined) lines.push(`    tr.font.name = ${JSON.stringify(s.font.name)};`)
        lines.push(`  }`)
      }

      return lines.join('\n')
    })
    .join('\n')
}

/**
 * Build the options fragment for insertSlidesFromBase64 (copy_slides). The
 * leading `, ` is included when any option is present so it can be appended
 * directly after the base64 argument. User-supplied targetSlideId/formatting
 * are JSON.stringify'd to stay inside their string literals.
 */
export function buildInsertOptions(formatting?: string, targetSlideId?: string): string {
  const optionsParts: string[] = []
  if (formatting) optionsParts.push(`formatting: ${JSON.stringify(formatting)}`)
  if (targetSlideId) optionsParts.push(`targetSlideId: ${JSON.stringify(targetSlideId)}`)
  return optionsParts.length > 0 ? `, { ${optionsParts.join(', ')} }` : ''
}

/**
 * Convert a glob pattern (only `*` wildcard) into a case-insensitive anchored
 * RegExp. Regex metacharacters in the pattern are escaped so a literal `.` or
 * `(` matches itself rather than acting as a regex operator.
 */
export function globToRegExp(pattern: string): RegExp {
  const escaped = pattern.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*')
  return new RegExp(`^${escaped}$`, 'i')
}

// ---------------------------------------------------------------------------
// Cross-group Office.js helpers
// ---------------------------------------------------------------------------

export async function getLocalCopyPath(
  connPool: ConnectionPool,
  target: { filePath: string | null; presentationId: string; ws: import('ws').WebSocket },
): Promise<string> {
  const filePath = target.filePath

  // Local file — already on disk
  if (filePath && !filePath.startsWith('http')) {
    if (!existsSync(filePath)) throw new Error(`Local file not found: ${filePath}`)
    return filePath
  }

  // Cloud file — check revision for cache validity
  const revCode = `
    var p = context.presentation.properties;
    p.load("revisionNumber");
    await context.sync();
    return p.revisionNumber;
  `
  const currentRevision = (await connPool.sendCommand('executeCode', { code: revCode }, target.ws)) as number

  const cached = localCopyCache.get(target.presentationId)
  if (cached && cached.revision === currentRevision && existsSync(cached.localPath)) {
    return cached.localPath
  }

  // Export fresh copy via Common API getFileAsync
  const exportCode = `
    return new Promise(function(resolve, reject) {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error.message));
          return;
        }
        var file = result.value;
        var sliceCount = file.sliceCount;
        var sliceData = [];
        var totalSize = 0;
        function getNextSlice(index) {
          if (index >= sliceCount) {
            file.closeAsync();
            var combined = new Uint8Array(totalSize);
            var offset = 0;
            for (var i = 0; i < sliceData.length; i++) {
              var arr = new Uint8Array(sliceData[i]);
              combined.set(arr, offset);
              offset += arr.length;
            }
            var binary = '';
            var chunk = 8192;
            for (var j = 0; j < combined.length; j += chunk) {
              binary += String.fromCharCode.apply(null, combined.subarray(j, Math.min(j + chunk, combined.length)));
            }
            resolve(btoa(binary));
            return;
          }
          file.getSliceAsync(index, function(sliceResult) {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(sliceResult.error.message));
              return;
            }
            sliceData.push(sliceResult.value.data);
            totalSize += sliceResult.value.data.length;
            getNextSlice(index + 1);
          });
        }
        getNextSlice(0);
      });
    });
  `
  const base64 = (await connPool.sendCommand('executeCode', { code: exportCode }, target.ws, 120_000)) as string

  const filename = filePath ? decodeURIComponent(filePath.split('/').pop() || 'presentation.pptx') : 'presentation.pptx'
  const dest = join(tmpdir(), `pptbridge-${Date.now()}-${filename}`)
  writeFileSync(dest, Buffer.from(base64, 'base64'))

  localCopyCache.set(target.presentationId, { localPath: dest, revision: currentRevision })
  return dest
}

export interface LayoutUsage {
  name: string
  id: string
  usedBySlides: number[]
}
export async function getLayoutUsage(connPool: ConnectionPool, ws: import('ws').WebSocket): Promise<LayoutUsage[]> {
  const code = `
    var slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    for (var i = 0; i < slides.items.length; i++) {
      slides.items[i].layout.load("name,id");
    }
    await context.sync();
    var seen = {};
    var layouts = [];
    for (var i = 0; i < slides.items.length; i++) {
      var l = slides.items[i].layout;
      if (!seen[l.id]) {
        seen[l.id] = true;
        layouts.push({ name: l.name, id: l.id, usedBySlides: [i] });
      } else {
        for (var j = 0; j < layouts.length; j++) {
          if (layouts[j].id === l.id) { layouts[j].usedBySlides.push(i); break; }
        }
      }
    }
    return layouts;
  `
  return (await connPool.sendCommand('executeCode', { code }, ws)) as LayoutUsage[]
}
