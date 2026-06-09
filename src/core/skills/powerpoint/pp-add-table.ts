/**
 * PowerPoint `ppt-add-table` skill — add a real Table (grid of
 * cells) to the active slide. Distinguish from `ppt-add-text`
 * (single text box) and from a 2-D text-box grid (we use the
 * Office.js Table shape so PowerPoint recognises it as a table,
 * not a collection of text frames).
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  slideIndex: z.number().int().min(1).default(1),
  rows: z.number().int().min(1).max(50),
  columns: z.number().int().min(1).max(20),
  /** Optional 2-D array of cell text matching the row × column shape. */
  data: z.array(z.array(z.string())).optional(),
  /** Position by anchor: full-width strip at top / middle / bottom. */
  position: z.enum(['top', 'middle', 'bottom']).default('middle'),
})

type PptAddTableArgs = z.infer<typeof args>
type PptAddTableResult = { shapeId: string; slideIndex: number }

export const pptAddTable: Skill<PptAddTableArgs, PptAddTableResult> = {
  name: 'ppt-add-table',
  description:
    'Add a real table (grid of cells) to a slide. Optionally pre-fill rows/columns of text.',
  host: ['powerpoint'],
  args,

  async execute(input, _ctx): Promise<PptAddTableResult> {
    const ppt = (globalThis as any).PowerPoint
    if (!ppt) {
      throw new Error('PowerPoint namespace not available — sideload.')
    }
    return (await ppt.run(async (context: any) => {
      const slide = context.presentation.slides.getItemAt(input.slideIndex - 1)
      const slideW = context.presentation.pageWidth ?? slide.layout?.width ?? 960
      const slideH = context.presentation.pageHeight ?? slide.layout?.height ?? 540

      const tableHeight = Math.round(slideH / 3)
      const yByPos = {
        top: Math.round(slideH * 0.1),
        middle: Math.round((slideH - tableHeight) / 2),
        bottom: Math.round(slideH * 0.9 - tableHeight),
      } as const

      const table = slide.shapes.addTable(
        input.rows,
        input.columns,
        0,
        yByPos[input.position],
        slideW,
        tableHeight,
      )
      if (input.data) {
        for (let r = 0; r < input.data.length; r++) {
          for (let c = 0; c < (input.data[r]?.length ?? 0); c++) {
            table.getCellOrNullObject(r, c).text = input.data[r]![c]!
          }
        }
      }
      await context.sync()
      return { shapeId: String(table.id), slideIndex: input.slideIndex }
    })) as PptAddTableResult
  },
}
