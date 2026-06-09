/**
 * PowerPoint `ppt-add-image` skill — embed a base64 image into
 * the active slide. Keeps the LLM out of file-system APIs (we
 * accept image bytes as base64 inline).
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  slideIndex: z.number().int().min(1).default(1),
  /** Base64-encoded image bytes (PNG / JPEG / GIF / SVG). */
  base64: z.string().min(1),
  /** MIME type — drives the correct addImage overload. */
  mimeType: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/svg+xml']).default('image/png'),
  /** Position anchor. */
  position: z.enum(['top-left', 'top-right', 'center', 'bottom-left', 'bottom-right']).default('center'),
  /** Optional width override (in points). Defaults to half the slide width. */
  width: z.number().int().min(20).max(2000).optional(),
  /** Optional height override. */
  height: z.number().int().min(20).max(2000).optional(),
})

type PptAddImageArgs = z.infer<typeof args>
type PptAddImageResult = { shapeId: string; slideIndex: number }

export const pptAddImage: Skill<PptAddImageArgs, PptAddImageResult> = {
  name: 'ppt-add-image',
  description: 'Embed a base64-encoded image (PNG / JPEG / GIF / SVG) into a slide.',
  host: ['powerpoint'],
  args,

  async execute(input, _ctx): Promise<PptAddImageResult> {
    const ppt = (globalThis as any).PowerPoint
    if (!ppt) {
      throw new Error('PowerPoint namespace not available — sideload.')
    }
    return (await ppt.run(async (context: any) => {
      const slide = context.presentation.slides.getItemAt(input.slideIndex - 1)
      const slideW = context.presentation.pageWidth ?? slide.layout?.width ?? 960
      const slideH = context.presentation.pageHeight ?? slide.layout?.height ?? 540
      const w = input.width ?? Math.round(slideW / 2)
      const h = input.height ?? Math.round(slideH / 2)
      const leftByPos = {
        'top-left': Math.round(slideW * 0.05),
        'top-right': slideW - w - Math.round(slideW * 0.05),
        center: Math.round((slideW - w) / 2),
        'bottom-left': Math.round(slideW * 0.05),
        'bottom-right': slideW - w - Math.round(slideW * 0.05),
      } as const
      const topByPos = {
        'top-left': Math.round(slideH * 0.05),
        'top-right': Math.round(slideH * 0.05),
        center: Math.round((slideH - h) / 2),
        'bottom-left': slideH - h - Math.round(slideH * 0.05),
        'bottom-right': slideH - h - Math.round(slideH * 0.05),
      } as const
      const img = slide.shapes.addImage(input.base64, input.mimeType)
      img.left = leftByPos[input.position]
      img.top = topByPos[input.position]
      img.width = w
      img.height = h
      await context.sync()
      return { shapeId: String(img.id), slideIndex: input.slideIndex }
    })) as PptAddImageResult
  },
}
