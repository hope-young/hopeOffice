/**
 * Excel `add-table` skill — a proper Excel Table (ListObject)
 * created from a range. Returns the table id + name so follow-up
 * skills can target it.
 *
 * Note: this is a *structured* Excel Table (with headers + filter
 * buttons + auto-expanding range), NOT a plain range formatted as
 * a grid. Distinct from the Word `add-table` which inserts a
 * real Word table.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  sheetName: z.string().default('Sheet1'),
  /** Range address like 'A1:B5'. The first row is treated as the
   *  header. */
  dataRange: z.string().min(1),
  /** Display name. Must be unique within the workbook. */
  tableName: z.string().default('Table1'),
  /** Optional style preset. Office.js exposes ~12; we expose the
   *  four most common. */
  style: z
    .enum(['light', 'medium', 'dark', 'none'])
    .default('light'),
})

type AddTableArgs = z.infer<typeof args>
type AddTableResult = { tableId: string; tableName: string }

export const addTable: Skill<AddTableArgs, AddTableResult> = {
  name: 'xl-add-table',
  description:
    'Create a structured Excel Table (ListObject) from an existing range. Returns the table id and name.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<AddTableResult> {
    const ExcelNS = (globalThis as any).Excel
    if (!ExcelNS) {
      throw new Error(
        'Excel namespace not available — sideload into Excel first.',
      )
    }
    return (await ExcelNS.run(async (context: any) => {
      const sheet =
        context.workbook.worksheets.getItemOrNullObject(input.sheetName)
      const range = sheet.getRange(input.dataRange)
      const table = sheet.tables.add(range, /* hasHeaders */ true)
      table.name = input.tableName
      if (input.style !== 'none') {
        // Office.js string-literal style names.
        const styleMap = {
          light: 'TableStyleLight1',
          medium: 'TableStyleMedium2',
          dark: 'TableStyleDark1',
        } as const
        table.style = styleMap[input.style]
      }
      await context.sync()
      return { tableId: String(table.id), tableName: input.tableName }
    })) as AddTableResult
  },
}
