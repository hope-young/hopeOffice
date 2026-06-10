/**
 * Excel `format-chart` skill.
 *
 * Tweaks the visual style of an existing chart: line colour,
 * line weight, and line style. The chart is referenced by the
 * `chartId` returned from a previous `add-chart` call. With no
 * `chartId` argument we operate on the first chart of the
 * active sheet, which is what the LLM usually wants when the
 * user says "the chart in the spreadsheet".
 *
 * Without this skill the LLM could only create charts (the
 * older `add-chart`) but not restyle them, which left SCI-paper
 * style requests stuck on a default-coloured line and no way
 * to fix it from chat.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const HEX_RE = /^#?([0-9a-fA-F]{6})$/

const args = z.object({
  chartId: z.string().optional(),
  sheetName: z.string().default('Sheet1'),
  /** Hex colour, e.g. "#1f77b4" or "1f77b4". */
  lineColor: z
    .string()
    .regex(HEX_RE, 'lineColor must be a 6-digit hex code (with or without #)')
    .optional(),
  /** Line weight in points. Excel clamps to [0.25, 12]. */
  lineWeight: z.number().min(0.25).max(12).optional(),
  /**
   * One of the Office.js `Excel.ChartLineStyle` values. We
   * accept the user-friendly names and map them to the SDK
   * constants inside the skill so the LLM can stay readable.
   */
  lineStyle: z
    .enum([
      'solid',
      'dash',
      'dot',
      'dash-dot',
      'dash-dot-dot',
      'none',
    ])
    .optional(),
})

type FormatChartArgs = z.infer<typeof args>
type FormatChartResult = { applied: string[] }

export const formatChart: Skill<FormatChartArgs, FormatChartResult> = {
  name: 'format-chart',
  description:
    'Restyle an existing Excel chart: line colour (hex), line weight (pt), and line style (solid/dash/dot/…). Reference the chart by the id returned from add-chart, or omit to target the first chart on the sheet.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<FormatChartResult> {
    type ExcelGlobal = { run: (cb: (ctx: Excel.RequestContext) => Promise<unknown>) => Promise<unknown> }
    const excelNs = (globalThis as { Excel?: ExcelGlobal }).Excel
    if (!excelNs) {
      throw new Error(
        'Excel namespace not available — this skill only runs inside an Excel task pane.',
      )
    }

    const lineStyleMap = {
      solid: 'Solid',
      dash: 'Dash',
      dot: 'Dot',
      'dash-dot': 'DashDot',
      'dash-dot-dot': 'DashDotDot',
      none: 'None',
    } as const

    return await excelNs.run(async (context: Excel.RequestContext) => {
      const sheet = context.workbook.worksheets.getItem(input.sheetName)
      // ChartCollection has no `getFirstOrNullObject` in Excel's
      // current typings, so we load the whole items array and
      // pick the first one if no id was provided.
      const charts = sheet.charts
      charts.load('items/id')
      let chart: Excel.Chart | null = null
      await context.sync()
      if (input.chartId) {
        chart = sheet.charts.getItemOrNullObject(input.chartId)
        chart.load('id')
        await context.sync()
      } else if (charts.items.length > 0) {
        chart = charts.items[0]
      }
      if (chart && !chart.isNullObject) {
        // Line properties live on each *series*, not the chart
        // itself (Office.js has no `chart.format.line`). We loop
        // every series so multi-line charts get the new style
        // applied uniformly. The runtime cost is trivial —
        // sync batches the writes.
        const series = chart.series
        series.load('items')
        await context.sync()
        for (let i = 0; i < series.items.length; i++) {
          const s = series.items[i]
          s.load(['format'])
          await context.sync()
          if (input.lineColor) {
            const hex = input.lineColor.replace(/^#/, '').toUpperCase()
            s.format.line.color = `#${hex}`
          }
          if (input.lineWeight !== undefined) {
            s.format.line.weight = input.lineWeight
          }
          if (input.lineStyle) {
            // The `style` setter exists on the runtime
            // ChartLineFormat object (Excel writes it back to
            // the workbook just fine) but the typings shipped
            // in @types/office-js only expose `color` and
            // `weight`. Cast through `unknown` to keep the
            // build green without lying to the rest of the
            // code about the shape of this object.
            ;(s.format.line as unknown as {
              style: Excel.ChartLineStyle
            }).style = lineStyleMap[
              input.lineStyle
            ] as unknown as Excel.ChartLineStyle
          }
        }
      }
      await context.sync()
      return {
        applied: [
          input.lineColor ? `color=${input.lineColor}` : null,
          input.lineWeight !== undefined ? `weight=${input.lineWeight}pt` : null,
          input.lineStyle ? `style=${input.lineStyle}` : null,
        ].filter((s): s is string => s !== null),
      }
    }) as FormatChartResult
  },
}
