/**
 * Word `set-font` skill — set font properties (name, size, bold /
 * italic / underline, color) on a range of text. Either an
 * explicit range (by id from a prior find) or "the whole body".
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** Either a rangeId from a prior find-text, or 'all' to apply
   *  to the entire body. */
  target: z.union([z.string(), z.literal('all')]).default('all'),
  fontName: z.string().optional(),
  fontSize: z.number().int().min(1).max(400).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.enum(['none', 'single', 'double', 'dotted', 'wavy']).optional(),
  /** Hex color (e.g. '#ff0000' or 'ff0000'). Office.js expects
   *  '#RRGGBB'. */
  color: z.string().regex(/^#?[0-9a-fA-F]{6}$/).optional(),
})

type SetFontArgs = z.infer<typeof args>
type SetFontResult = { applied: boolean }

export const setFont: Skill<SetFontArgs, SetFontResult> = {
  name: 'set-font',
  description:
    'Set font name, size, weight, style, and color on a range of text. Pass a rangeId from find-text, or "all" for the whole body.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<SetFontResult> {
    const wordNs = (globalThis as any).Word
    if (!wordNs) {
      throw new Error('Word namespace not available — sideload.')
    }
    return (await wordNs.run(async (context: any) => {
      // Phase 8 scope keeps this simple — no range lookup; we
      // accept 'all' only. Specific range support (a Phase 9
      // polish) would track a rangeId-to-Range map at the
      // orchestrator level.
      if (input.target !== 'all') {
        throw new Error(
          'set-font: per-range targeting is Phase 9. Use target: "all" for now.',
        )
      }
      const range = context.document.body.getRange()
      const font = range.font
      if (input.fontName) font.name = input.fontName
      if (input.fontSize) font.size = input.fontSize
      if (input.bold !== undefined) font.bold = input.bold
      if (input.italic !== undefined) font.italic = input.italic
      if (input.underline) font.underline = input.underline
      if (input.color) {
        const hex = input.color.replace('#', '').toUpperCase()
        font.color = `#${hex}`
      }
      await context.sync()
      return { applied: true }
    })) as SetFontResult
  },
}
