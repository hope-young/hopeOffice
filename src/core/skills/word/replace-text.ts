/**
 * Word `replace-text` skill — find and replace text across the
 * document body. Case-sensitive by default; pass `matchCase: true`
 * for whole-word. Uses `range.replace` under the hood.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** Substring to find. */
  needle: z.string().min(1),
  /** Replacement text. Empty string deletes occurrences. */
  replacement: z.string(),
  matchCase: z.boolean().default(true),
  /** Replace all occurrences vs only the first. */
  replaceAll: z.boolean().default(true),
})

type ReplaceTextArgs = z.infer<typeof args>
type ReplaceTextResult = { replacements: number }

export const replaceText: Skill<ReplaceTextArgs, ReplaceTextResult> = {
  name: 'replace-text',
  description:
    'Find and replace text across the document body. Empty replacement deletes occurrences.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<ReplaceTextResult> {
    const wordNs = (globalThis as any).Word
    if (!wordNs) {
      throw new Error('Word namespace not available — sideload.')
    }
    return (await wordNs.run(async (context: any) => {
      const results = context.document.body.search(input.needle, {
        matchCase: input.matchCase,
      })
      results.load('items')
      await context.sync()
      const items = results.items ?? []
      // Word's `range.replace` works on a single Range. We loop
      // and count successful replaces.
      let n = 0
      for (const r of items) {
        r.insertText(input.replacement, 'Replace')
        n++
        if (!input.replaceAll) break
      }
      await context.sync()
      return { replacements: n }
    })) as ReplaceTextResult
  },
}
