/**
 * Excel `set-chart-axis-title` skill.
 *
 * Sets the title text of one of the chart's two axes. Excel
 * distinguishes the value axis (Y) from the category axis (X)
 * — `valueAxis` is the vertical one with numbers, `categoryAxis`
 * the horizontal one with labels.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  chartId: z.string().optional(),
  sheetName: z.string().default('Sheet1'),
  /** Which axis to label. */
  axis: z.enum(['x', 'y']).default('y'),
  title: z.string().default(''),
  visible: z.boolean().default(true),
})

type SetChartAxisTitleArgs = z.infer<typeof args>
type SetChartAxisTitleResult = { axis: 'x' | 'y'; title: string }

export const setChartAxisTitle: Skill<SetChartAxisTitleArgs, SetChartAxisTitleResult> = {
  name: 'set-chart-axis-title',
  description:
    'Set the title of an Excel chart axis (X = categoryAxis / horizontal, Y = valueAxis / vertical). Pass empty title to clear.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SetChartAxisTitleResult> {
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
        const axis =
          input.axis === 'y' ? chart.axes.valueAxis : chart.axes.categoryAxis
        axis.title.text = input.title
        axis.title.visible = input.visible
      }
      await context.sync()
      return { axis: input.axis, title: input.title }
    }) as SetChartAxisTitleResult
  },
}
