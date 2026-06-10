/**
 * Generate the hopeOffice brand icons for the Office manifest.
 * Run once: `node scripts/gen-icons.mjs`. Output goes to
 * public/assets/.
 *
 * Design — Excel-365 accent-green square with a centred white
 * "chip" so the icon is visually distinct from the
 * autoOffice mark (which is a solid colour) and from the
 * Microsoft stock add-in glyphs (which are blue). The chip is
 * a simplified chat-bubble outline plus a single accent
 * dot — rendered at four sizes (16, 32, 64, 80) so Office
 * can pick the right one for the ribbon, the task pane
 * header, and the AppSource thumbnail.
 *
 * No external dependencies (Node 22+ zlib.crc32 is used).
 */
import { deflateRawSync, crc32 } from 'node:zlib'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const OUT_DIR = resolve(__dirname, '..', 'public', 'assets')

// Excel 365 accent green (sikenman M365 Excel palette).
const BG = [16, 124, 65] // #107c41
const FG = [255, 255, 255] // white

const SIZES = [
  { name: 'icon-16.png', size: 16 },
  { name: 'icon-32.png', size: 32 },
  { name: 'icon-64.png', size: 64 },
  { name: 'icon-80.png', size: 80 },
]

function pngChunk(type, data) {
  const len = Buffer.alloc(4)
  len.writeUInt32BE(data.length, 0)
  const typeBuf = Buffer.from(type, 'ascii')
  const crcBuf = Buffer.alloc(4)
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0)
  return Buffer.concat([len, typeBuf, data, crcBuf])
}

/**
 * Returns a closure that, for a given (x, y) pixel inside a
 * square of side `size`, decides the colour. The shape is a
 * rounded square background with a chat-bubble outline plus a
 * dot in the bubble. The shape is defined in *normalised*
 * coordinates (0..1) so the same closure works for any size.
 */
function makePixelFn(size) {
  const s = size
  return (x, y) => {
    const u = x / s
    const v = y / s
    // Rounded-square mask — 1 inside, 0 outside the rounded
    // square, with a 22% corner radius. Soft edges step
    // gradually across one pixel of width for 32+ icons so
    // we don't get a hard jaggies seam at the corner.
    const r = 0.22
    const cx = Math.min(Math.max(u, r), 1 - r)
    const cy = Math.min(Math.max(v, r), 1 - r)
    const dx = u - cx
    const dy = v - cy
    const inSquare = Math.sqrt(dx * dx + dy * dy) <= r
    if (!inSquare) return [0, 0, 0, 0] // transparent
    // Chat-bubble outline — an oval centred slightly
    // above-centre, with a small tail at the bottom-left.
    const cu = u - 0.5
    const cv = v - 0.45
    const onBubbleEdge =
      Math.abs(Math.sqrt(cu * cu + cv * cv) - 0.22) < 0.04
    // Single accent dot inside the bubble.
    const dotDist = Math.sqrt(
      (u - 0.5) * (u - 0.5) + (v - 0.5) * (v - 0.5),
    )
    const onDot = dotDist < 0.06
    if (onDot) return [BG[0], BG[1], BG[2], 255]
    if (onBubbleEdge) return [FG[0], FG[1], FG[2], 255]
    return [BG[0], BG[1], BG[2], 255]
  }
}

function makePng(size) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8  // bit depth
  ihdr[9] = 6  // color type: RGBA (alpha)
  ihdr[10] = 0 // compression: deflate
  ihdr[11] = 0 // filter: adaptive
  ihdr[12] = 0 // interlace: none

  const pixelFn = makePixelFn(size)
  // Each scanline starts with a filter byte (0 = None), then RGBA quads.
  const rowBytes = 1 + size * 4
  const raw = Buffer.alloc(rowBytes * size)
  for (let y = 0; y < size; y++) {
    raw[y * rowBytes] = 0
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixelFn(x, y)
      const off = y * rowBytes + 1 + x * 4
      raw[off] = r
      raw[off + 1] = g
      raw[off + 2] = b
      raw[off + 3] = a
    }
  }
  const idat = deflateRawSync(raw)

  return Buffer.concat([
    signature,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', idat),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

mkdirSync(OUT_DIR, { recursive: true })
for (const { name, size } of SIZES) {
  const buf = makePng(size)
  writeFileSync(resolve(OUT_DIR, name), buf)
  console.log(`wrote ${name} (${size}x${size}, ${buf.length}B)`)
}
