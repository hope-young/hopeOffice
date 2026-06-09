/**
 * Excel `sort-range` skill — sort a range by a single key column.
 * For multi-key sorts use `apply-sort` (Phase 8+); for now we
 * keep it to the common case the LLM reaches for first.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  sheetName: z.string().default('Sheet1'),
  range: z.string().min(1), // e.g. 'A1:D10'
  /** 1-based index of the key column within the range. */
  keyColumn: z.number().int().min(1).default(1),
  /** Sort direction. */
  ascending: z.boolean().default(true),
  /** Whether the first row of the range is a header row (left
   *  out of the sort). */
  hasHeader: z.boolean().default(true),
})

type SortRangeArgs = z.infer<typeof args>
type SortRangeResult = { range: string }

export const sortRange: Skill<SortRangeArgs, SortRangeResult> = {
  name: 'sort-range',
  description:
    'Sort a range by a single key column. The key column index is 1-based within the range.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<SortRangeResult> {
    const ExcelNS = (globalThis as any).Excel
    if (!ExcelNS) {
      throw new Error('Excel namespace not available — sideload into Excel.')
    }
    return (await ExcelNS.run(async (context: any) => {
      const sheet =
        context.workbook.worksheets.getItemOrNullObject(input.sheetName)
      const range = sheet.getRange(input.range)
      // SortField on the SortFields collection; the key column
      // here is `range.getCellAt(0, keyColumn - 1).entireColumn`
      // so the sort spans the whole range by the chosen column.
      const keyRange = range.getCell(0, input.keyColumn - 1).entireColumn
      range.sort.apply([
        {
          key: keyRange,
          ascending: input.ascending,
        },
      ])
      await context.sync()
      return { range: input.range }
    })) as SortRangeResult
  },
}
