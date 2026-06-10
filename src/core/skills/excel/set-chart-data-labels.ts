/**
 * Excel `set-chart-data-labels` skill.
 *
 * Toggles the per-point data labels on a chart and picks
 * where they sit relative to the marker. The LLM almost
 * always wants the default ("show value, position outside end")
 * for SCI-paper-style line charts, so the defaults are tuned
 * to that case.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  chartId: z.string().optional(),
  sheetName: z.string().default('Sheet1'),
  showValue: z.boolean().default(true),
  /** Office.js ChartDataLabelPosition. */
  position: z
    .enum([
      'outside-end',
      'inside-end',
      'center',
      'inside-base',
      'best-fit',
    ])
    .default('outside-end'),
})

type SetChartDataLabelsArgs = z.infer<typeof args>
type SetChartDataLabelsResult = { showValue: boolean; position: string }

const POSITION_MAP = {
  'outside-end': 'OutsideEnd',
  'inside-end': 'InsideEnd',
  center: 'Center',
  'inside-base': 'InsideBase',
  'best-fit': 'BestFit',
} as const

export const setChartDataLabels: Skill<SetChartDataLabelsArgs, SetChartDataLabelsResult> = {
  name: 'set-chart-data-labels',
  description:
    'Toggle data labels on an Excel chart and choose their position. Defaults to show-value + outside-end, the SCI-paper look.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SetChartDataLabelsResult> {
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
        chart.dataLabels.showValue = input.showValue
        chart.dataLabels.position =
          POSITION_MAP[input.position] as unknown as Excel.ChartDataLabelPosition
      }
      await context.sync()
      return { showValue: input.showValue, position: input.position }
    }) as SetChartDataLabelsResult
  },
}
