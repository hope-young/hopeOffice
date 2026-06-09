/**
 * Word `add-text` skill.
 *
 * Inserts a paragraph at the start / end of the document body, or
 * after the current selection. Returns the new paragraph's id so
 * follow-up skills can reference it (format, attach a comment,
 * etc.).
 *
 * Mirrors `powerpoint/add-text` in shape so the LLM only has to
 * learn one mental model across the two text-host skills.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** The text to insert. */
  text: z.string().min(1),
  /** Where to put the new paragraph. */
  position: z.enum(['start', 'end', 'cursor']).default('cursor'),
  /** Optional font-size in points. */
  fontSize: z.number().int().min(8).max(144).optional(),
  /** Bold the inserted paragraph. */
  bold: z.boolean().default(false),
})

type AddTextArgs = z.infer<typeof args>
type AddTextResult = { paragraphId: string }

export const addText: Skill<AddTextArgs, AddTextResult> = {
  // Host prefix matches `ppt-add-text` — keeps the LLM's tool
  // palette unambiguous when it can see both hosts' skill
  // descriptors (system prompt + Zod schemas).
  name: 'word-add-text',
  description:
    'Insert a paragraph of text into the active Word document. Optionally position it at the start / end of the body, or right after the current selection.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<AddTextResult> {
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
      let paragraph: any
      if (input.position === 'start') {
        paragraph = context.document.body.insertParagraph(input.text, 'Start')
      } else if (input.position === 'cursor') {
        const sel = context.document.getSelection()
        paragraph = sel
          ? sel.insertParagraph(input.text, 'After')
          : context.document.body.insertParagraph(input.text, 'End')
      } else {
        paragraph = context.document.body.insertParagraph(input.text, 'End')
      }
      if (input.bold) paragraph.font.bold = true
      if (input.fontSize) paragraph.font.size = input.fontSize
      await context.sync()
      return { paragraphId: String(paragraph.id) }
    })) as AddTextResult
  },
}
