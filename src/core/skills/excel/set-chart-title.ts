/**
 * Excel `set-chart-title` skill.
 *
 * Sets or clears a chart's title. The LLM can also flip the
 * title's visibility flag independently — useful for charts
 * that already have a text set but the user wants them
 * hidden for a screenshot.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  chartId: z.string().optional(),
  sheetName: z.string().default('Sheet1'),
  /** New title text. Pass empty string to clear. */
  title: z.string().default(''),
  visible: z.boolean().default(true),
})

type SetChartTitleArgs = z.infer<typeof args>
type SetChartTitleResult = { title: string; visible: boolean }

export const setChartTitle: Skill<SetChartTitleArgs, SetChartTitleResult> = {
  name: 'set-chart-title',
  description:
    'Set or clear an Excel chart title, and toggle its visibility. Pass an empty title to clear.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SetChartTitleResult> {
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
        chart.title.text = input.title
        chart.title.visible = input.visible
      }
      await context.sync()
      return { title: input.title, visible: input.visible }
    }) as SetChartTitleResult
  },
}
