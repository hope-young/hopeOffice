/**
 * Excel `set-cell-value` skill — set a literal value in one or
 * more cells. For formulas use `add-formula`; for ranges with
 * mixed types, batch in Zod.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  sheetName: z.string().default('Sheet1'),
  range: z.string().min(1), // e.g. 'A1' or 'A1:B2'
  /** 2D array of values matching the range shape. The LLM is
   *  expected to read the range size first (or trust the user
   *  to have told them) and pass the matching matrix. */
  values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
})

type SetCellValueArgs = z.infer<typeof args>
type SetCellValueResult = { range: string; cellsWritten: number }

export const setCellValue: Skill<SetCellValueArgs, SetCellValueResult> = {
  name: 'set-cell-value',
  description:
    'Set literal values in a range. Pass a 2D array of strings / numbers / booleans matching the range shape.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SetCellValueResult> {
    const ExcelNS = (globalThis as any).Excel
    if (!ExcelNS) {
      throw new Error('Excel namespace not available — sideload into Excel.')
    }
    return (await ExcelNS.run(async (context: any) => {
      const sheet =
        context.workbook.worksheets.getItemOrNullObject(input.sheetName)
      const range = sheet.getRange(input.range)
      range.values = input.values
      await context.sync()
      const cells = input.values.reduce(
        (n: number, row: unknown[]) => n + row.length,
        0,
      )
      return { range: input.range, cellsWritten: cells }
    })) as SetCellValueResult
  },
}
