/**
 * Generate minimal placeholder PNG icons for the Office manifest.
 * Run once: `node scripts/gen-icons.mjs`. Output goes to public/assets/.
 *
 * Solid 4-channel blue squares — Office will scale them on the ribbon.
 * No external dependencies (Node 22+ zlib.crc32 is used).
 */
import { deflateRawSync, crc32 } from 'node:zlib'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const OUT_DIR = resolve(__dirname, '..', 'public', 'assets')

// Hope-Office brand blue (matches the placeholder theme).
const COLOR = [37, 99, 235]

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

function makePng(size, [r, g, b]) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8  // bit depth
  ihdr[9] = 2  // color type: RGB
  ihdr[10] = 0 // compression: deflate
  ihdr[11] = 0 // filter: adaptive
  ihdr[12] = 0 // interlace: none

  // Each scanline starts with a filter byte (0 = None), then RGB triples.
  const rowBytes = 1 + size * 3
  const raw = Buffer.alloc(rowBytes * size)
  for (let y = 0; y < size; y++) {
    raw[y * rowBytes] = 0
    for (let x = 0; x < size; x++) {
      const off = y * rowBytes + 1 + x * 3
      raw[off] = r
      raw[off + 1] = g
      raw[off + 2] = b
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
  const buf = makePng(size, COLOR)
  writeFileSync(resolve(OUT_DIR, name), buf)
  console.log(`wrote ${name} (${size}x${size}, ${buf.length}B)`)
}
