/**
 * Word `add-image` skill — embed a base64 image at the end of the
 * document body (inline shape). Mirrors the PowerPoint variant
 * in shape but uses Word's `insertInlinePictureFromBase64`.
 */
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  base64: z.string().min(1),
  mimeType: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/svg+xml']).default('image/png'),
  /** Optional width override. Height is auto-scaled to preserve
   *  aspect ratio unless `height` is also given. */
  width: z.number().int().min(20).max(2000).optional(),
  height: z.number().int().min(20).max(2000).optional(),
})

type AddImageArgs = z.infer<typeof args>
type AddImageResult = { shapeId: string }

export const addImage: Skill<AddImageArgs, AddImageResult> = {
  name: 'add-image',
  description: 'Embed a base64 image at the end of the document body.',
  host: ['word'],
  args,

  async execute(input, _ctx): Promise<AddImageResult> {
    const wordNs = (globalThis as any).Word
    if (!wordNs) {
      throw new Error('Word namespace not available — sideload.')
    }
    return (await wordNs.run(async (context: any) => {
      const pic = context.document.body.insertInlinePictureFromBase64(
        input.base64,
        input.mimeType,
        'End',
      )
      if (input.width) pic.width = input.width
      if (input.height) pic.height = input.height
      await context.sync()
      return { shapeId: String(pic.id) }
    })) as AddImageResult
  },
}
