/**
 * PowerPoint `ppt-add-bullets` skill — add a bullet-list text
 * shape to a slide. Each item becomes one paragraph in a single
 * text frame.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  slideIndex: z.number().int().min(1).default(1),
  items: z.array(z.string().min(1)).min(1).max(20),
  position: z.enum(['top', 'middle', 'bottom']).default('middle'),
  fontSize: z.number().int().min(8).max(144).optional(),
})

type PptAddBulletsArgs = z.infer<typeof args>
type PptAddBulletsResult = { shapeId: string; slideIndex: number }

export const pptAddBullets: Skill<PptAddBulletsArgs, PptAddBulletsResult> = {
  name: 'ppt-add-bullets',
  description: 'Add a bullet-list text frame to a slide. Each item is one bullet.',
  host: ['powerpoint'],
  args,

  async execute(input, _ctx): Promise<PptAddBulletsResult> {
    const ppt = (globalThis as any).PowerPoint
    if (!ppt) {
      throw new Error('PowerPoint namespace not available — sideload.')
    }
    return (await ppt.run(async (context: any) => {
      const slide = context.presentation.slides.getItemAt(input.slideIndex - 1)
      const slideW = context.presentation.pageWidth ?? slide.layout?.width ?? 960
      const slideH = context.presentation.pageHeight ?? slide.layout?.height ?? 540
      const boxH = Math.round(slideH / 3)
      const yByPos = {
        top: Math.round(slideH * 0.1),
        middle: Math.round((slideH - boxH) / 2),
        bottom: Math.round(slideH * 0.9 - boxH),
      } as const
      const body = input.items.map((line) => `• ${line}`).join('\n')
      const tb = slide.shapes.addTextBox(body)
      tb.left = Math.round(slideW * 0.1)
      tb.top = yByPos[input.position]
      tb.width = Math.round(slideW * 0.8)
      tb.height = boxH
      if (input.fontSize) {
        tb.textFrame.textRange.font.size = input.fontSize
      }
      await context.sync()
      return { shapeId: String(tb.id), slideIndex: input.slideIndex }
    })) as PptAddBulletsResult
  },
}
