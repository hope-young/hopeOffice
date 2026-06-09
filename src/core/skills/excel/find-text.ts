/**
 * Excel `find-text` skill — return addresses of all cells in a
 * sheet matching a substring (case-insensitive). Useful for the
 * LLM to answer "where is X?" questions and to drive follow-up
 * edit skills.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  sheetName: z.string().default('Sheet1'),
  /** Substring to search for. Case-insensitive. */
  needle: z.string().min(1),
  /** Optional cap on the number of hits returned. */
  limit: z.number().int().min(1).max(500).default(50),
})

type FindTextArgs = z.infer<typeof args>
type FindTextHit = { address: string; value: string }
type FindTextResult = { hits: FindTextHit[] }

export const findText: Skill<FindTextArgs, FindTextResult> = {
  name: 'xl-find-text',
  description:
    'Find all cells in the active sheet whose value contains the given substring (case-insensitive). Returns the addresses.',
  host: ['excel'],
  args,

  async execute(input, _ctx): Promise<FindTextResult> {
    const ExcelNS = (globalThis as any).Excel
    if (!ExcelNS) {
      throw new Error('Excel namespace not available — sideload into Excel.')
    }
    return (await ExcelNS.run(async (context: any) => {
      const sheet =
        context.workbook.worksheets.getItemOrNullObject(input.sheetName)
      const found = sheet.findAll(input.needle, {
        completeMatch: false,
        matchCase: false,
      })
      const hits: FindTextHit[] = []
      const range = found.getFirst()
      // Walk the found collection via load on each — office.js
      // returns a RangeAreas object; we just pull address + value.
      found.load('items')
      await context.sync()
      const items = (found.items ?? []).slice(0, input.limit)
      for (const r of items) {
        r.load(['address', 'values'])
      }
      await context.sync()
      for (const r of items) {
        const v = r.values?.[0]?.[0]
        hits.push({ address: r.address, value: v == null ? '' : String(v) })
      }
      // range is unused at runtime; reference so TS doesn't trip
      // the unused-binding check.
      void range
      return { hits }
    })) as FindTextResult
  },
}
