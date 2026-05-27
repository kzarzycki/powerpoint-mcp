import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { buildChartRelationship, buildChartXml, buildGraphicFrame, resolveChartPosition } from '../chart-builder.ts'
import { withTool } from '../tool-helpers.ts'
import { applyZipEditAndReimport, exportSlide, extractZipFiles, listZipPaths } from '../xml-helpers.ts'
import { getConcurrentWarning } from './shared.ts'

export function registerChartTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'edit_slide_chart',
    'Create a chart on a slide from structured data. Generates all OOXML automatically (chart XML, rels, graphic frame, Content_Types). Supports column, bar, line, pie, doughnut, and area charts with multiple series.',
    {
      slideIndex: z.number().int().min(0).describe('Zero-based slide index'),
      chartType: z.enum(['column', 'bar', 'line', 'pie', 'doughnut', 'area']).describe('Chart type'),
      title: z.string().describe('Chart title'),
      categories: z.array(z.string()).describe('Category labels (x-axis or pie slices)'),
      series: z
        .array(
          z.object({
            name: z.string().describe('Series name'),
            values: z.array(z.number()).describe('Data values (one per category)'),
          }),
        )
        .min(1)
        .describe('Data series'),
      position: z
        .object({
          left: z.number().optional().describe('Left position in points'),
          top: z.number().optional().describe('Top position in points'),
          width: z.number().optional().describe('Width in points'),
          height: z.number().optional().describe('Height in points'),
        })
        .optional()
        .describe('Chart position in points. Defaults to centered on slide.'),
      options: z
        .object({
          stacked: z.boolean().optional().describe('Use stacked grouping (bar/column/line/area)'),
          showDataLabels: z.boolean().optional().describe('Show data labels (default true)'),
          showLegend: z.boolean().optional().describe('Show legend (default true)'),
          legendPosition: z
            .enum(['t', 'b', 'l', 'r'])
            .optional()
            .describe('Legend position: t=top, b=bottom, l=left, r=right'),
        })
        .optional()
        .describe('Chart options'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slideIndex, chartType, title, categories, series, position, options, presentationId }) => {
      const target = pool.resolveTarget(presentationId)

      // 1. Export the slide
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { zip } = await extractZipFiles(exported.base64)

      // 2. Determine next chart number (scan existing ppt/charts/ in zip)
      const existingPaths = listZipPaths(zip)
      const chartPaths = existingPaths.filter((p) => p.startsWith('ppt/charts/chart') && p.endsWith('.xml'))
      const chartNums = chartPaths.map((p) => {
        const m = p.match(/chart(\d+)\.xml$/)
        return m ? Number(m[1]) : 0
      })
      const nextChartNum = chartNums.length > 0 ? Math.max(...chartNums) + 1 : 1
      const chartFileName = `chart${nextChartNum}.xml`
      const chartZipPath = `ppt/charts/${chartFileName}`

      // 3. Determine next rId from slide rels
      const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
      const relsContent = zip.file(relsPath)
        ? await zip.file(relsPath)!.async('string')
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
      const rIdMatches = [...relsContent.matchAll(/Id="rId(\d+)"/g)]
      const rIdNums = rIdMatches.map((m) => Number(m[1]))
      const nextRIdNum = rIdNums.length > 0 ? Math.max(...rIdNums) + 1 : 1
      const rId = `rId${nextRIdNum}`

      // 4. Generate chart XML
      const chartXml = buildChartXml(chartType, title, categories, series, options)

      // 5. Generate graphic frame and inject into slide XML
      const slideXmlPath = 'ppt/slides/slide1.xml'
      const slideXml = await zip.file(slideXmlPath)!.async('string')
      const emuPos = resolveChartPosition(position)

      // Find the highest shape ID in the slide to avoid conflicts
      const shapeIdMatches = [...slideXml.matchAll(/id="(\d+)"/g)]
      const shapeIds = shapeIdMatches.map((m) => Number(m[1]))
      const nextShapeId = shapeIds.length > 0 ? Math.max(...shapeIds) + 1 : 100
      const chartName = `Chart ${nextChartNum}`

      const graphicFrame = buildGraphicFrame(rId, emuPos, chartName, nextShapeId)

      // Inject graphic frame before </p:spTree>
      const modifiedSlideXml = slideXml.replace('</p:spTree>', `${graphicFrame}</p:spTree>`)

      // 6. Add chart relationship to rels
      const relEntry = buildChartRelationship(rId, `../charts/${chartFileName}`)
      const modifiedRels = relsContent.replace('</Relationships>', `${relEntry}</Relationships>`)

      // 7. Write all files and reimport
      const files: Record<string, string> = {
        [slideXmlPath]: modifiedSlideXml,
        [chartZipPath]: chartXml,
        [relsPath]: modifiedRels,
      }

      await applyZipEditAndReimport(pool, exported, files, target.ws)

      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text =
        JSON.stringify(
          {
            success: true,
            chartType,
            title,
            seriesCount: series.length,
            categoryCount: categories.length,
            chartFile: chartZipPath,
          },
          null,
          2,
        ) + (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )
}
