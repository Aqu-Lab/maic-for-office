/** Dev-only: inspect image blocks of a docx to find non-decodable dataUrls. */
import { readFileSync } from 'node:fs'
import { parseDocx } from '@maic/docx-engine'

const path = process.argv[2]
const bytes = new Uint8Array(readFileSync(path))
const doc = await parseDocx(bytes)

let i = 0
for (const block of doc.blocks) {
  const entries = []
  if (block.type === 'image' && block.imageDataUrl) {
    entries.push({ where: 'block', url: block.imageDataUrl, w: block.imageWidthPx, h: block.imageHeightPx, wrap: block.imageWrap })
  }
  for (const run of block.runs ?? []) {
    if (run.image?.dataUrl) {
      entries.push({ where: 'run', url: run.image.dataUrl, w: run.image.widthPx, h: run.image.heightPx, wrap: run.image.wrap })
    }
  }
  for (const e of entries) {
    i++
    const mime = /^data:([^;,]+)/.exec(e.url)?.[1] ?? '?'
    console.log(`#${i} ${e.where} mime=${mime} ${e.w ?? '?'}x${e.h ?? '?'} wrap=${e.wrap ?? '-'} len=${e.url.length}`)
  }
}
console.log(`total images: ${i}`)
