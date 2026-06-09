/**
 * Word `add-table` skill. Source of truth: SPEC_DETAILS §6.
 *
 * Inserts an N×M table at the start / end of the document body,
 * or at the current cursor position. Optionally pre-fills cells
 * (in which case the first row is treated as the header — no
 * bold-styling is applied; the LLM has already styled the data
 * as it likes). When `data` is absent and `hasHeader` is true,
 * the first row gets bolded so it visually stands out.
 *
 * Sandbox + dev-browser notes: same as the PowerPoint / Excel
 * skills — the Word namespace is injected by Office at task-pane
 * load; dev preview without Office surfaces a clean error via
 * the tool-call-error event path.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  rows: z.number().int().min(1).max(50),
  columns: z.number().int().min(1).max(20),
  /** Optional pre-fill; when supplied, the first row is taken to
   *  be the header text (no auto-bolding — caller already styled it). */
  data: z.array(z.array(z.string())).optional(),
  hasHeader: z.boolean().default(true),
  position: z
    .enum(['start', 'end', 'cursor'])
    .default('end'),
})

type AddTableArgs = z.infer<typeof args>
type AddTableResult = { tableId: string; rows: number; columns: number }

export const addTable: Skill<AddTableArgs, AddTableResult> = {
  name: 'add-table',
  description:
    'Insert a table into the active Word document. Optionally pre-fill cells. Returns the table id so follow-up skills can reference it.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<AddTableResult> {
    type WordGlobal = {
      run: (cb: (ctx: unknown) => Promise<unknown>) => Promise<unknown>
    }
    const wordNs = (globalThis as { Word?: WordGlobal }).Word
    if (!wordNs) {
      throw new Error(
        'Word namespace not available — this skill only runs inside a Word task pane. ' +
          'Sideload the add-in into Word and try again.',
      )
    }

    return (await wordNs.run(async (context: any) => {
      const body = context.document.body
      // The reference location string Word expects:
      //  - 'Start' / 'End' for body-relative insertion
      //  - For 'cursor' we read the current selection and insert
      //    after it; fall back to End when there's no selection.
      let refLocation: string
      if (input.position === 'start') {
        refLocation = 'Start'
      } else if (input.position === 'cursor') {
        const sel = context.document.getSelection()
        refLocation = sel ? 'After' : 'End'
      } else {
        refLocation = 'End'
      }

      // Word's TableCollection.insertTable(rowCount, columnCount, location, values)
      const table = body.insertTable(
        input.rows,
        input.columns,
        refLocation,
        input.data ?? [],
      )
      // Built-in table style — string literal (the enum exists
      // but stringly-typed values are more stable across Word
      // versions than the TS enum names).
      table.styleBuiltIn = 'WellLightShadingAccent1'
      if (input.hasHeader && !input.data) {
        // Bold the first row to flag it as the header.
        table.rows.getFirst().font.bold = true
      }
      await context.sync()
      return {
        tableId: String(table.id),
        rows: input.rows,
        columns: input.columns,
      }
    })) as AddTableResult
  },
}
