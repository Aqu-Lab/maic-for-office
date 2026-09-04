/**
 * TIFF → PNG transcoding (browser side). Chromium cannot decode TIFF, so pictures
 * embedded as ppt/media/*.tif(f) need UTIF decode + canvas re-encode for display.
 * Adapted from the slides app's main-process version: pngjs encoding is replaced
 * with an OffscreenCanvas/canvas draw, which the Obsidian renderer always has.
 */
import UTIF from 'utif2'

/** PNG data URL of the largest TIFF page, or null when undecodable. */
export function tiffToPngDataUrl(bytes: Uint8Array): string | null {
  try {
    const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const ifds = UTIF.decode(buf)
    if (!ifds.length) return null
    // Multi-page/multi-resolution TIFFs: pick the largest page
    let page = ifds[0]!
    for (const ifd of ifds) {
      UTIF.decodeImage(buf, ifd)
      const cur = (ifd.width || 0) * (ifd.height || 0)
      if (cur > (page.width || 0) * (page.height || 0)) page = ifd
    }
    const width = page.width
    const height = page.height
    if (!width || !height) return null
    const rgba = UTIF.toRGBA8(page)
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    const imageData = ctx.createImageData(width, height)
    imageData.data.set(rgba)
    ctx.putImageData(imageData, 0, 0)
    return canvas.toDataURL('image/png')
  } catch {
    return null
  }
}
