/**
 * Word `find-text` skill — find all occurrences of a substring in
 * the body and return their (rangeId, surrounding text) for
 * downstream edits. Same shape as the Excel variant but bound to
 * Word's search results collection.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  needle: z.string().min(1),
  matchCase: z.boolean().default(true),
  limit: z.number().int().min(1).max(500).default(50),
})

type FindTextArgs = z.infer<typeof args>
type FindTextHit = { rangeId: string; context: string }
type FindTextResult = { hits: FindTextHit[] }

export const findText: Skill<FindTextArgs, FindTextResult> = {
  name: 'word-find-text',
  description:
    'Find all occurrences of a substring in the document body. Returns rangeIds (so follow-up skills can edit them) plus a context snippet.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<FindTextResult> {
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
      const items = (results.items ?? []).slice(0, input.limit)
      for (const r of items) {
        r.load(['text', 'parentBody'])
      }
      await context.sync()
      const hits: FindTextHit[] = []
      for (const r of items) {
        // `parentBody.getText()` is heavy; just take a short
        // window around the hit so the LLM has context.
        const fullText = r.parentBody?.text ?? ''
        const idx = fullText.indexOf(input.needle)
        const start = Math.max(0, idx - 20)
        const end = Math.min(fullText.length, idx + input.needle.length + 20)
        const context = idx >= 0 ? fullText.slice(start, end) : input.needle
        hits.push({ rangeId: String(r.id), context })
      }
      return { hits }
    })) as FindTextResult
  },
}
