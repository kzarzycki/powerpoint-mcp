import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import { getConcurrentWarning } from './shared.ts'

export function registerExecTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'execute_officejs',
    "Execute arbitrary Office.js code inside the live PowerPoint presentation. The code runs inside PowerPoint.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. For positioning, all values are in points (1 point = 1/72 inch). Common operations: add shapes, set text, change colors, add/delete slides.",
    {
      code: z
        .string()
        .describe(
          "Office.js code to execute. Runs inside PowerPoint.run() with 'context' available. Use 'return' to send back a result.",
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ code, presentationId }) => {
      const target = pool.resolveTarget(presentationId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text = JSON.stringify(result ?? { success: true }, null, 2) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'verify_slides',
    'Run programmatic checks on a slide: detect overlapping shapes, out-of-bounds shapes, empty text, tiny shapes, unused placeholders, placeholder drift from layout defaults, and full-bleed background covers. Returns a list of issues found. Uses the same shape data as inspect_slide — no OOXML needed.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      checks: z
        .array(
          z.enum([
            'overlap',
            'bounds',
            'empty_text',
            'tiny_shapes',
            'unused_placeholder',
            'layout_drift',
            'background_cover',
          ]),
        )
        .optional()
        .describe('Checks to run. Default: all checks.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideIndex, checks, presentationId }) => {
      const target = pool.resolveTarget(presentationId)
      const enabledChecks = checks ?? [
        'overlap',
        'bounds',
        'empty_text',
        'tiny_shapes',
        'unused_placeholder',
        'layout_drift',
        'background_cover',
      ]

      const checkLayoutDrift = enabledChecks.includes('layout_drift')

      // Reuse inspect_slide's shape-loading logic
      const code = `
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if (${slideIndex} >= slides.items.length) {
            throw new Error("Slide index " + ${slideIndex} + " out of range (presentation has " + slides.items.length + " slides)");
          }
          var slide = slides.items[${slideIndex}];
          slide.shapes.load("items");
          await context.sync();
          var shapes = [];
          var placeholderTypes = [];
          for (var i = 0; i < slide.shapes.items.length; i++) {
            var s = slide.shapes.items[i];
            var info = {
              name: s.name,
              id: s.id,
              left: s.left,
              top: s.top,
              width: s.width,
              height: s.height
            };
            try {
              var tf = s.getTextFrameOrNullObject();
              tf.load(["hasText", "textRange"]);
              await context.sync();
              if (!tf.isNullObject) {
                info.text = tf.hasText ? tf.textRange.text : "";
                info.hasText = tf.hasText;
              }
            } catch (e) {}
            if (s.type === "Placeholder") {
              try {
                var pf = s.placeholderFormat;
                pf.load("type");
                await context.sync();
                info.isPlaceholder = true;
                info.placeholderType = pf.type;
                placeholderTypes.push({ shapeIndex: i, type: pf.type });
              } catch (e) {}
            }
            shapes.push(info);
          }

          // Conditionally load layout placeholder positions for drift check
          var layoutMap = {};
          if (${checkLayoutDrift} && placeholderTypes.length > 0) {
            try {
              var layout = slide.layout;
              layout.load("name");
              var layoutShapes = layout.shapes;
              layoutShapes.load("items");
              await context.sync();
              for (var li = 0; li < layoutShapes.items.length; li++) {
                var ls = layoutShapes.items[li];
                if (ls.type !== "Placeholder") continue;
                try {
                  var lph = ls.placeholderFormat;
                  lph.load("type");
                  ls.load("left,top,width,height,name");
                  await context.sync();
                  layoutMap[lph.type] = {
                    name: ls.name,
                    left: ls.left,
                    top: ls.top,
                    width: ls.width,
                    height: ls.height
                  };
                } catch (e) {}
              }
            } catch (e) {}
            // Attach layout match to shapes
            for (var pi = 0; pi < placeholderTypes.length; pi++) {
              var pt = placeholderTypes[pi];
              var match = layoutMap[pt.type];
              if (match) {
                shapes[pt.shapeIndex].layoutMatch = match;
              }
            }
          }

          // Also get slide dimensions
          var ps = context.presentation.pageSetup;
          ps.load("slideWidth,slideHeight");
          await context.sync();
          return { shapes: shapes, slideWidth: ps.slideWidth, slideHeight: ps.slideHeight };
        `
      const slideData = (await pool.sendCommand('executeCode', { code }, target.ws)) as {
        shapes: Array<{
          name: string
          id: string
          left: number
          top: number
          width: number
          height: number
          text?: string
          hasText?: boolean
          isPlaceholder?: boolean
          placeholderType?: string
          layoutMatch?: { name: string; left: number; top: number; width: number; height: number }
        }>
        slideWidth: number
        slideHeight: number
      }

      const issues: Array<{
        type: string
        severity: 'warning' | 'error'
        shapeIds: string[]
        description: string
      }> = []

      const { shapes, slideWidth, slideHeight } = slideData

      // Overlap check: AABB collision
      if (enabledChecks.includes('overlap')) {
        for (let i = 0; i < shapes.length; i++) {
          for (let j = i + 1; j < shapes.length; j++) {
            const a = shapes[i]!
            const b = shapes[j]!
            if (
              a.left < b.left + b.width &&
              a.left + a.width > b.left &&
              a.top < b.top + b.height &&
              a.top + a.height > b.top
            ) {
              issues.push({
                type: 'overlap',
                severity: 'warning',
                shapeIds: [a.id, b.id],
                description: `"${a.name}" and "${b.name}" overlap`,
              })
            }
          }
        }
      }

      // Bounds check: shape extends beyond slide
      if (enabledChecks.includes('bounds')) {
        for (const s of shapes) {
          const outOfBounds: string[] = []
          if (s.left < 0) outOfBounds.push('left of slide')
          if (s.top < 0) outOfBounds.push('above slide')
          if (s.left + s.width > slideWidth) outOfBounds.push('right of slide')
          if (s.top + s.height > slideHeight) outOfBounds.push('below slide')
          if (outOfBounds.length > 0) {
            issues.push({
              type: 'bounds',
              severity: 'warning',
              shapeIds: [s.id],
              description: `"${s.name}" extends ${outOfBounds.join(', ')}`,
            })
          }
        }
      }

      // Empty text check
      if (enabledChecks.includes('empty_text')) {
        for (const s of shapes) {
          if (s.text !== undefined && s.text.trim() === '') {
            issues.push({
              type: 'empty_text',
              severity: 'warning',
              shapeIds: [s.id],
              description: `"${s.name}" has an empty text frame`,
            })
          }
        }
      }

      // Tiny shapes check
      if (enabledChecks.includes('tiny_shapes')) {
        for (const s of shapes) {
          if (s.width < 10 || s.height < 10) {
            issues.push({
              type: 'tiny_shapes',
              severity: 'warning',
              shapeIds: [s.id],
              description: `"${s.name}" is very small (${s.width.toFixed(1)} x ${s.height.toFixed(1)} pt)`,
            })
          }
        }
      }

      // Unused placeholder check: placeholder shapes with no text content
      if (enabledChecks.includes('unused_placeholder')) {
        for (const s of shapes) {
          if (s.isPlaceholder && !s.hasText) {
            issues.push({
              type: 'unused_placeholder',
              severity: 'warning',
              shapeIds: [s.id],
              description: `"${s.name}" is an unused placeholder — delete it or fill it with content`,
            })
          }
        }
      }

      // Layout drift check: placeholder position vs layout default
      if (checkLayoutDrift) {
        const DRIFT_THRESHOLD = 2 // points
        for (const s of shapes) {
          if (!s.isPlaceholder || !s.layoutMatch) continue
          const lm = s.layoutMatch
          const drifts: string[] = []
          if (Math.abs(s.left - lm.left) > DRIFT_THRESHOLD) drifts.push(`left: ${s.left} vs layout ${lm.left}`)
          if (Math.abs(s.top - lm.top) > DRIFT_THRESHOLD) drifts.push(`top: ${s.top} vs layout ${lm.top}`)
          if (Math.abs(s.width - lm.width) > DRIFT_THRESHOLD) drifts.push(`width: ${s.width} vs layout ${lm.width}`)
          if (Math.abs(s.height - lm.height) > DRIFT_THRESHOLD)
            drifts.push(`height: ${s.height} vs layout ${lm.height}`)
          if (drifts.length > 0) {
            issues.push({
              type: 'layout_drift',
              severity: 'warning',
              shapeIds: [s.id],
              description: `"${s.name}" drifted from layout: ${drifts.join(', ')}`,
            })
          }
        }
      }

      // Background cover check: non-placeholder shapes covering most of the slide
      if (enabledChecks.includes('background_cover')) {
        const dimThreshold = 0.85
        const areaThreshold = 0.9
        const slideArea = slideWidth * slideHeight
        for (const s of shapes) {
          if (s.isPlaceholder) continue
          const widthRatio = s.width / slideWidth
          const heightRatio = s.height / slideHeight
          if (
            widthRatio >= dimThreshold &&
            heightRatio >= dimThreshold &&
            s.width * s.height >= slideArea * areaThreshold
          ) {
            issues.push({
              type: 'background_cover',
              severity: 'error',
              shapeIds: [s.id],
              description: `"${s.name}" covers ${(widthRatio * 100).toFixed(0)}% x ${(heightRatio * 100).toFixed(0)}% of the slide — this destroys the layout background, logo, and design system. Delete this shape and use the layout's background instead.`,
            })
          }
        }
      }

      const result = { slideIndex, shapeCount: shapes.length, issueCount: issues.length, issues }
      return { content: [{ type: 'text' as const, text: JSON.stringify(result, null, 2) }] }
    }),
  )
}
