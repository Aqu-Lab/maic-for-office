import { TFile, WorkspaceLeaf } from 'obsidian'
import { openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide, HeuristicMetrics, type RenderSlide } from '@genoffice/pptx-render'
import { BaseOfficeView } from './base-view'
import { makeMediaResolver } from './render/pptx/media-resolver'
import { mountPptxReader } from './pptx-canvas'
import type OfficeViewerPlugin from '../main'

export const VIEW_TYPE_PPTX = 'office-viewer-pptx'

/** Render-tree width all pages are built at; DOM/CSS scales the stage for display. */
const FIT_WIDTH_PX = 1280

/**
 * Read-only .pptx viewer. Parsing and render-tree building run in-process with the
 * same engines the Slides app uses; drawing is the app's static Konva layer.
 *
 * Fonts: text metrics use HeuristicMetrics (no font file access) in this version.
 * System-font metrics via opentype.js (apps/slides/src/main/fonts.ts) is the
 * next fidelity step — see README.
 */
export class PptxView extends BaseOfficeView {
  private unmountReader: (() => void) | null = null

  constructor(leaf: WorkspaceLeaf, plugin: OfficeViewerPlugin) {
    super(leaf, plugin)
  }

  getViewType(): string {
    return VIEW_TYPE_PPTX
  }

  getDisplayText(): string {
    const file = this.getFile()
    return file ? file.name : 'Presentation'
  }

  getIcon(): string {
    return 'presentation'
  }

  protected extension(): string {
    return 'pptx'
  }

  protected onCleanup(): void {
    this.unmountReader?.()
    this.unmountReader = null
  }

  protected async renderInto(container: HTMLElement, bytes: Uint8Array, _file: TFile): Promise<void> {
    const opened: OpenedPptx = await openPptx(bytes)
    const metrics = new HeuristicMetrics()
    const slides: RenderSlide[] = opened.deck.slides.map((s, i) =>
      buildRenderSlide(s, opened.deck.size, {
        fitWidthPx: FIT_WIDTH_PX,
        media: makeMediaResolver(opened, s.path),
        metrics,
        slideNo: i + 1,
      }),
    )
    container.empty()
    const host = container.createDiv('ovx-pptx-host')
    host.style.height = '100%'
    this.unmountReader = mountPptxReader(host, slides, this.plugin.settings.pptxFitWidth)
  }
}
