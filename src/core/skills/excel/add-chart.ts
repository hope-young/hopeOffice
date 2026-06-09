/**
 * Excel `add-chart` skill.
 *
 * Inserts a chart object into the active Excel worksheet. The
 * returned `chartId` lets follow-up skills (resize, format, move
 * to a slide, …) reference the chart later in the conversation.
 *
 * Skill body runs in the task-pane context, where Office.js injects
 * the `Excel` global. In dev-mode browser preview (no Office) the
 * skill throws at the top, and the orchestrator surfaces that as a
 * `tool-call-error` event so the LLM sees a clean message instead
 * of a JS exception leaking into the chat.
 *
 * Sandbox execution (per SPEC §2) is a Phase 6 concern; for now the
 * trust model is "LLM picks a tool, Zod validates, the skill calls
 * office.js directly". The pattern is identical to the
 * SPEC_DETAILS §6 add-table example, just on Excel's namespace.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** Worksheet to insert the chart into. Defaults to the active sheet. */
  sheetName: z.string().default('Sheet1'),
  /** Chart type. Office.js exposes 70+ specific enum values
   *  (ColumnClustered, Line, Pie, …). We pick four common ones
   *  the LLM is most likely to ask for, normalised to the
   *  string the user actually meant. */
  chartType: z
    .enum(['line', 'bar-clustered', 'pie', 'column-clustered'])
    .default('column-clustered'),
  /** Range address like 'A1:B5'. Drives the chart's data. */
  dataRange: z.string().min(1),
  /** Optional chart title. */
  title: z.string().optional(),
})

type AddChartArgs = z.infer<typeof args>
type AddChartResult = { chartId: string; title?: string }

export const addChart: Skill<AddChartArgs, AddChartResult> = {
  name: 'add-chart',
  description:
    'Insert a chart into the active Excel worksheet. Returns the chart id so follow-up skills can reference it.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<AddChartResult> {
    // The Excel namespace is injected into the task pane by Office
    // on load. We look it up via globalThis so the file type-checks
    // even when the @types/office-js typings are not in scope.
    type ExcelGlobal = { run: (callback: (ctx: Excel.RequestContext) => Promise<unknown>) => Promise<unknown> }
    const excelNs = (globalThis as { Excel?: ExcelGlobal }).Excel
    if (!excelNs) {
      throw new Error(
        'Excel namespace not available — this skill only runs inside an Excel task pane. ' +
          'Sideload the add-in into Excel and try again.',
      )
    }

    return await excelNs.run(async (context: Excel.RequestContext) => {
      const sheet = context.workbook.worksheets.getItem(input.sheetName)
      const range = sheet.getRange(input.dataRange)
      // Map our friendly enum to the literal Office.js ChartType
      // values (`Excel.ChartType` is a TS string-literal union of
      // 70+ members — we accept four and assert the cast).
      const chartType = (
        {
          line: 'Line',
          'bar-clustered': 'BarClustered',
          pie: 'Pie',
          'column-clustered': 'ColumnClustered',
        } as const
      )[input.chartType] as 'Line'
      const chart = sheet.charts.add(chartType, range, 'Auto')
      if (input.title) chart.title.text = input.title
      await context.sync()
      return { chartId: String(chart.id), title: input.title }
    }) as AddChartResult
  },
}
