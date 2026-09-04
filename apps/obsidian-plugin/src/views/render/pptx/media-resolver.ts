/**
 * Image mediaRef → dataUrl for buildRenderSlide's `media` option.
 *
 * Adapted from the slides app's session-state makeMediaResolver: TIFF transcodes to
 * PNG (Chromium can't decode it), JPEG EXIF orientation is neutralized so rotated
 * pixels don't double-rotate, SVG passes through (the theme-slot retint of Office
 * SVG exports is an editor-only nicety, skipped here), everything else becomes a
 * plain base64 data URL.
 */
import type { OpenedPptx } from '@maic/pptx-engine'
import { displayMime } from './media-mime'
import { neutralizeJpegOrientation } from './jpeg-orientation'
import { tiffToPngDataUrl } from './tiff-decode'

export function makeMediaResolver(opened: OpenedPptx, _slidePath?: string) {
  const cache = new Map<string, string | undefined>()
  return (mediaRef: string): string | undefined => {
    if (cache.has(mediaRef)) return cache.get(mediaRef)
    const bytes = opened.archive.readBytes(mediaRef)
    let url: string | undefined
    if (bytes) {
      const mime = displayMime(mediaRef, bytes)
      if (mime === 'image/tiff') {
        url = tiffToPngDataUrl(bytes) ?? undefined
      } else if (mime === 'image/svg+xml') {
        url = bytesToDataUrl(bytes, mime, 'utf8')
      } else {
        // PowerPoint ignores EXIF orientation; Chromium applies it on decode — neutralize
        // the flag so rotated-pixel JPEGs with a shape-level rot don't double-rotate
        const served = mime === 'image/jpeg' ? neutralizeJpegOrientation(bytes) : bytes
        url = bytesToDataUrl(served, mime)
      }
    }
    cache.set(mediaRef, url)
    return url
  }
}

function bytesToDataUrl(bytes: Uint8Array, mime: string, encoding: 'binary' | 'utf8' = 'binary'): string {
  let binary = ''
  const chunk = 0x8000
  if (encoding === 'utf8') {
    const text = new TextDecoder().decode(bytes)
    for (let i = 0; i < text.length; i += chunk) {
      binary += String.fromCharCode(...new TextEncoder().encode(text.slice(i, i + chunk)))
    }
  } else {
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
    }
  }
  return `data:${mime};base64,${btoa(binary)}`
}
