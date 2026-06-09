/**
 * Excel `add-formula` skill — set a formula in a single cell.
 * Distinguishes from `set-cell-value` (literal value) by accepting
 * a `=`-prefixed string that Excel evaluates.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  sheetName: z.string().default('Sheet1'),
  cell: z.string().min(1), // e.g. 'A1'
  formula: z.string().min(1), // e.g. '=SUM(B1:B10)'
})

type AddFormulaArgs = z.infer<typeof args>
type AddFormulaResult = { cell: string }

export const addFormula: Skill<AddFormulaArgs, AddFormulaResult> = {
  name: 'add-formula',
  description:
    'Set a formula in a single cell. The formula must start with `=`.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<AddFormulaResult> {
    const ExcelNS = (globalThis as any).Excel
    if (!ExcelNS) {
      throw new Error('Excel namespace not available — sideload into Excel.')
    }
    if (!input.formula.startsWith('=')) {
      throw new Error('Formula must start with `=` (got a literal value?)')
    }
    return (await ExcelNS.run(async (context: any) => {
      const sheet =
        context.workbook.worksheets.getItemOrNullObject(input.sheetName)
      const range = sheet.getRange(input.cell)
      range.formulas = [[input.formula]]
      await context.sync()
      return { cell: input.cell }
    })) as AddFormulaResult
  },
}
