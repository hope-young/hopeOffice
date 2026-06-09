/**
 * Word `add-page-break` skill — insert a hard page break at the
 * end of the document body, or right after the current selection.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  position: z.enum(['end', 'cursor']).default('end'),
})

type AddPageBreakArgs = z.infer<typeof args>
type AddPageBreakResult = { rangeId: string }

export const addPageBreak: Skill<AddPageBreakArgs, AddPageBreakResult> = {
  name: 'add-page-break',
  description:
    'Insert a hard page break at the end of the document body, or right after the current selection.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<AddPageBreakResult> {
    const wordNs = (globalThis as any).Word
    if (!wordNs) {
      throw new Error('Word namespace not available — sideload.')
    }
    return (await wordNs.run(async (context: any) => {
      const target =
        input.position === 'cursor'
          ? (context.document.getSelection() ?? context.document.body)
          : context.document.body
      const range = target.insertBreak('Page', 'After')
      await context.sync()
      return { rangeId: String(range.id) }
    })) as AddPageBreakResult
  },
}
