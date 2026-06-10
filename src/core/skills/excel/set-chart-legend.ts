/**
 * Excel `set-chart-legend` skill.
 *
 * Toggles the legend on a chart and chooses its anchor.
 * Pass visible: false to hide, or pick one of the four
 * Office.js positions.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  chartId: z.string().optional(),
  sheetName: z.string().default('Sheet1'),
  visible: z.boolean().default(true),
  position: z
    .enum(['top', 'bottom', 'left', 'right', 'corner'])
    .default('bottom'),
})

type SetChartLegendArgs = z.infer<typeof args>
type SetChartLegendResult = { visible: boolean; position: string }

const POSITION_MAP = {
  top: 'Top',
  bottom: 'Bottom',
  left: 'Left',
  right: 'Right',
  corner: 'Corner',
} as const

export const setChartLegend: Skill<SetChartLegendArgs, SetChartLegendResult> = {
  name: 'set-chart-legend',
  description: 'Show, hide, or reposition the legend of an Excel chart.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SetChartLegendResult> {
    type ExcelGlobal = { run: (cb: (ctx: Excel.RequestContext) => Promise<unknown>) => Promise<unknown> }
    const excelNs = (globalThis as { Excel?: ExcelGlobal }).Excel
    if (!excelNs) {
      throw new Error('Excel namespace not available.')
    }

    return await excelNs.run(async (context: Excel.RequestContext) => {
      const sheet = context.workbook.worksheets.getItem(input.sheetName)
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
        chart.legend.visible = input.visible
        if (input.visible) {
          chart.legend.position =
            POSITION_MAP[input.position] as unknown as Excel.ChartLegendPosition
        }
      }
      await context.sync()
      return { visible: input.visible, position: input.position }
    }) as SetChartLegendResult
  },
}
