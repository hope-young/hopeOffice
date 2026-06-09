/**
 * PowerPoint `add-text` skill.
 *
 * Inserts a text box on a given slide with user-controlled text,
 * vertical anchor (top / middle / bottom), and optional font size.
 * Returns the shape id so follow-up skills (move, format, link
 * to chart) can reference the text box.
 *
 * Sandbox + dev-browser notes: the PowerPoint namespace is
 * injected by Office at task-pane load. Dev preview without Office
 * surfaces a clean "namespace not available" error through the
 * tool-call-error path (verified end-to-end in Phase 5).
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  /** 1-indexed slide number. Users think 1-based; PowerPoint is 0-indexed
   *  internally — we translate below. */
  slideIndex: z.number().int().min(1).default(1),
  /** The text to put in the text box. */
  text: z.string().min(1),
  /** Vertical anchor. Width is full slide width; height derived from
   *  text length at the chosen font size. */
  position: z.enum(['top', 'middle', 'bottom']).default('middle'),
  /** Font size in points. */
  fontSize: z.number().int().min(8).max(144).optional(),
})

type AddTextArgs = z.infer<typeof args>
type AddTextResult = { shapeId: string; slideIndex: number }

export const addText: Skill<AddTextArgs, AddTextResult> = {
  name: 'add-text',
  description:
    'Insert a text box on a slide with the given text. Slide numbers are 1-based. Optionally set a vertical anchor and font size.',
  host: ['powerpoint'],

  args,

  async execute(input, _ctx): Promise<AddTextResult> {
    // `@types/office-js` doesn't ship PowerPoint types (the
    // community-maintained package only covers Word + Excel), so
    // we cast to `any` at the boundary. The runtime API is the
    // one documented at
    // https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/
    const ppt = (globalThis as { PowerPoint?: { run: (cb: (ctx: unknown) => Promise<unknown>) => Promise<unknown> } }).PowerPoint
    if (!ppt) {
      throw new Error(
        'PowerPoint namespace not available — this skill only runs inside a PowerPoint task pane. ' +
          'Sideload the add-in into PowerPoint and try again.',
      )
    }

    return (await ppt.run(async (context: any) => {
      const slides = context.presentation.slides
      const slide = slides.getItemAt(input.slideIndex - 1)
      // Slide width / height live on the slide's layout in
      // PowerPoint.js, not on the presentation object.
      const slideW: number = context.presentation.pageWidth ?? slide.layout?.width ?? 960
      const slideH: number = context.presentation.pageHeight ?? slide.layout?.height ?? 540

      // Text box sizing: full-width strip, 1/6 of slide height,
      // anchored top / middle / bottom. The exact height is a
      // guess that survives most slide layouts.
      const boxH = Math.round(slideH / 6)
      const yByPos = {
        top: Math.round(slideH * 0.05),
        middle: Math.round((slideH - boxH) / 2),
        bottom: Math.round(slideH * 0.95 - boxH),
      } as const
      const textBox = slide.shapes.addTextBox(input.text)
      // Position the inserted text box via its .top / .left / .width /
      // .height setters (the addTextBox one-arg overload doesn't take
      // a rect). We set them right after.
      textBox.left = 0
      textBox.top = yByPos[input.position]
      textBox.width = slideW
      textBox.height = boxH
      if (input.fontSize) {
        textBox.textFrame.textRange.font.size = input.fontSize
      }
      await context.sync()
      return { shapeId: String(textBox.id), slideIndex: input.slideIndex }
    })) as AddTextResult
  },
}
