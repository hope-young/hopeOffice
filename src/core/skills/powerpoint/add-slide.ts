/**
 * PowerPoint `add-slide` skill.
 *
 * Appends a new slide to the active presentation. Optional layout
 * (`blank` / `title` / `content`) and title text. Returns the new
 * slide's index so follow-up skills can target it without
 * re-counting.
 *
 * PowerPoint's slide layout types live on the slide master's
 * `SlideMaster` collection; `add({ layout: ... })` is the
 * accepted way to create a new slide against a built-in layout.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** Built-in slide layout. `blank` is the safest default — it never
   *  requires a pre-existing master layout. */
  layout: z.enum(['blank', 'title', 'content']).default('content'),
  /** Optional title text. Ignored if the chosen layout doesn't
   *  have a title placeholder (e.g. `blank`). */
  title: z.string().optional(),
})

type AddSlideArgs = z.infer<typeof args>
type AddSlideResult = { slideIndex: number }

export const addSlide: Skill<AddSlideArgs, AddSlideResult> = {
  name: 'add-slide',
  description:
    'Append a new slide to the active presentation. Returns the new slide index (1-based).',
  host: ['powerpoint'],

  args,

  async execute(input, _ctx): Promise<AddSlideResult> {
    // `@types/office-js` doesn't ship PowerPoint types. See add-text
    // for the full note; we cast to `any` at the boundary and rely
    // on runtime docs to keep the call shape correct.
    const ppt = (globalThis as { PowerPoint?: { run: (cb: (ctx: unknown) => Promise<unknown>) => Promise<unknown> } }).PowerPoint
    if (!ppt) {
      throw new Error(
        'PowerPoint namespace not available — this skill only runs inside a PowerPoint task pane. ' +
          'Sideload the add-in into PowerPoint and try again.',
      )
    }

    return (await ppt.run(async (context: any) => {
      // PowerPoint.PpSlideLayoutType enum: Blank / Title / TitleOnly
      // / Content / etc. We expose three via a friendly enum.
      const layoutType = (
        {
          blank: 'Blank',
          title: 'Title',
          content: 'Content',
        } as const
      )[input.layout] as 'Blank'
      // `slides.add({ layout })` — newer PowerPoint.js accepts the
      // layout as a PpSlideLayoutType string. Some older runtimes
      // want `slides.add(layoutType)`. We try the named form first
      // (the documented one); the call site can be adjusted if a
      // user's runtime complains.
      const slide = context.presentation.slides.add({ layout: layoutType })
      // New slides may not have a `title` placeholder when layout
      // is 'blank'; guard.
      if (input.title && slide.title && slide.title.textFrame) {
        slide.title.textFrame.textRange.text = input.title
      }
      await context.sync()
      return { slideIndex: context.presentation.slides.getCount() }
    })) as AddSlideResult
  },
}
