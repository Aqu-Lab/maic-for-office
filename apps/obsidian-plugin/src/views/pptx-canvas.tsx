/**
 * pptx read-only canvas — RenderSlide[] → Konva stages (React root inside the ItemView).
 *
 * Mirrors the slides app's SlideThumb static rendering: one Stage per page,
 * background rect + StaticNode per RenderTree node. A sidebar of scaled-down
 * pages drives the main stage (no virtualization — decks beyond ~100 pages
 * still work, pages render lazily as their thumbnails scroll into view).
 */
import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { Stage, Layer, Rect } from 'react-konva'
import type { RenderFill, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import { fillToKonva } from './render/pptx/konva-adapter'
import { StaticNode } from './render/pptx/NodeBody'
import { createImageLoader } from './render/pptx/image-loader'

function bgFillProps(slide: RenderSlide, images: Map<string, HTMLImageElement>) {
  const f = fillToKonva(slide.background, slide.widthPx, slide.heightPx, images)
  return Object.keys(f).length ? f : { fill: '#ffffff' }
}

/** One read-only page scaled to `width`. */
function SlidePage({
  slide,
  images,
  width,
  interactive,
}: {
  slide: RenderSlide
  images: Map<string, HTMLImageElement>
  width: number
  interactive?: boolean
}) {
  const scale = width / slide.widthPx
  const height = slide.heightPx * scale
  return (
    <Stage
      width={Math.round(width)}
      height={Math.round(height)}
      listening={interactive === true}
      style={{ pointerEvents: interactive ? 'auto' : 'none' }}
    >
      <Layer scaleX={scale} scaleY={scale} listening={false}>
        <Rect x={0} y={0} width={slide.widthPx} height={slide.heightPx} {...bgFillProps(slide, images)} />
        {slide.nodes.map((n) => (
          <StaticNode key={n.id} node={n} images={images} />
        ))}
      </Layer>
    </Stage>
  )
}

/** Collect picture urls across pages so every thumbnail can paint before being visited. */
function collectImageUrls(slides: RenderSlide[]): Set<string> {
  const urls = new Set<string>()
  const addFillUrl = (fill: RenderFill | undefined) => {
    if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
  }
  const walk = (nodes: readonly RenderNode[]) => {
    for (const n of nodes) {
      if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
      if ((n.type === 'shape' || n.type === 'text') && n.fill) addFillUrl(n.fill)
      if (n.type === 'chart') addFillUrl((n as { bgFill?: RenderFill }).bgFill)
      if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
      if (n.type === 'table' && Array.isArray(n.cells)) {
        for (const c of n.cells) if (c.fill) addFillUrl(c.fill)
      }
    }
  }
  for (const s of slides) {
    addFillUrl(s.background)
    walk(s.nodes)
  }
  return urls
}

function PptxReader({ slides, fitWidth }: { slides: RenderSlide[]; fitWidth: boolean }) {
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  const [current, setCurrent] = useState(0)
  const [mainWidth, setMainWidth] = useState(900)
  const mainHostRef = useRef<HTMLDivElement | null>(null)
  const loaderRef = useRef<ReturnType<typeof createImageLoader> | null>(null)

  const urls = useMemo(() => collectImageUrls(slides), [slides])
  useEffect(() => {
    if (!loaderRef.current) {
      loaderRef.current = createImageLoader((entries) => {
        setImages((prev) => {
          const m = new Map(prev)
          for (const [k, v] of entries) m.set(k, v)
          return m
        })
      })
    }
    loaderRef.current.load(urls)
    return () => {
      loaderRef.current?.dispose()
      loaderRef.current = null
    }
  }, [urls])

  // Track the main pane width so fit-to-width keeps working through pane resizes
  useEffect(() => {
    const host = mainHostRef.current
    if (!host) return
    const ro = new ResizeObserver(() => {
      setMainWidth(Math.max(host.clientWidth - 40, 320))
    })
    ro.observe(host)
    return () => ro.disconnect()
  }, [])

  const slide = slides[current] ?? slides[0]
  const thumbW = 132
  const mainW = fitWidth ? mainWidth : Math.min(mainWidth, slide?.widthPx ?? mainWidth)

  return (
    <div className="ovx-pptx">
      <div className="ovx-pptx-sidebar">
        {slides.map((s, i) => (
          <div
            key={i}
            className={`ovx-pptx-thumb${i === current ? ' ovx-active' : ''}`}
            onClick={() => setCurrent(i)}
          >
            <SlidePage slide={s} images={images} width={thumbW} />
            <span className="ovx-pptx-thumb-no">{i + 1}</span>
          </div>
        ))}
      </div>
      <div className="ovx-pptx-main" ref={mainHostRef}>
        {slide && (
          <div className="ovx-pptx-stage" style={{ width: Math.round(mainW) }}>
            <SlidePage slide={slide} images={images} width={mainW} interactive />
          </div>
        )}
      </div>
    </div>
  )
}

/** Mount a read-only pptx reader into `host`; returns the unmount function. */
export function mountPptxReader(host: HTMLElement, slides: RenderSlide[], fitWidth: boolean): () => void {
  const root = createRoot(host)
  root.render(<PptxReader slides={slides} fitWidth={fitWidth} />)
  return () => {
    root.unmount()
  }
}
