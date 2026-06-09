/**
 * PowerPoint `duplicate-slide` skill — clones a slide and inserts
 * the copy right after the source. Useful for "make N variants
 * of slide 3" workflows where the LLM then edits the duplicates.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** 1-based slide number of the slide to clone. */
  sourceSlide: z.number().int().min(1),
})

type DuplicateSlideArgs = z.infer<typeof args>
type DuplicateSlideResult = { newSlideIndex: number }

export const duplicateSlide: Skill<DuplicateSlideArgs, DuplicateSlideResult> = {
  name: 'duplicate-slide',
  description:
    'Duplicate a slide. The copy is inserted right after the source. Returns the new slide index (1-based).',
  host: ['powerpoint'],
  args,

  async execute(input, _ctx): Promise<DuplicateSlideResult> {
    const ppt = (globalThis as any).PowerPoint
    if (!ppt) {
      throw new Error('PowerPoint namespace not available — sideload.')
    }
    return (await ppt.run(async (context: any) => {
      const slide = context.presentation.slides.getItemAt(input.sourceSlide - 1)
      const copy = slide.duplicate()
      // `duplicate()` inserts right after the source; move it
      // explicitly to the same position just in case runtime
      // semantics differ.
      copy.moveTo(input.sourceSlide) // 1-based, after the source
      await context.sync()
      return { newSlideIndex: input.sourceSlide + 1 }
    })) as DuplicateSlideResult
  },
}
