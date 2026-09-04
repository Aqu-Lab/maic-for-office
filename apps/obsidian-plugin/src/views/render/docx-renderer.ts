/**
 * docx read-only renderer — ParsedDoc (block tree from @genoffice/docx-engine) → DOM.
 *
 * Scope: structured, Word-like preview. Paragraph/character styles (the parser already
 * flattens the basedOn chain into StyleInfo.display), lists with real numbering
 * definitions, tables (spans/merges/fills), inline + block images, hyperlinks,
 * header/footer strip, TOC lines and field results. No pagination measurement —
 * content flows down one continuous paper, like a print preview approximation.
 *
 * Uses only standard DOM APIs (no Obsidian element extensions) so the renderer is
 * testable under jsdom and independent of host prototype patches.
 */
import type {
  Block,
  DocDefaults,
  HfParagraph,
  ParsedDoc,
  ParaFormat,
  Run,
  StyleInfo,
  TableModel,
  TableLook,
  TableStyleDisplay,
} from '@genoffice/docx-engine'
import { metafileToDataUrl } from '@genoffice/docx-engine/metafile'

const PX_PER_TWIP = 96 / 1440
const PX_PER_HALF_POINT = 96 / 144

/** w:highlight named colors → hex (OOXML ST_HighlightColor). */
const HIGHLIGHT: Record<string, string> = {
  yellow: 'ffff00', green: '00ff00', cyan: '00ffff', magenta: 'ff00ff',
  blue: '0000ff', red: 'ff0000', darkBlue: '000080', darkCyan: '008080',
  darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
  darkYellow: '808000', darkGray: '808080', lightGray: 'c0c0c0',
  black: '000000', white: 'ffffff',
}

const ROMAN = ['', 'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x',
  'xi', 'xii', 'xiii', 'xiv', 'xv', 'xvi', 'xvii', 'xviii', 'xix', 'xx']

function toRoman(n: number, upper: boolean): string {
  const table = upper ? ROMAN.map((s) => s.toUpperCase()) : ROMAN
  return n >= 1 && n < table.length ? table[n]! : String(n)
}

/** Format one numbering counter per w:numFmt (the formats Word documents actually use). */
function formatCounter(numFmt: string, n: number): string {
  switch (numFmt) {
    case 'decimal': return String(n)
    case 'lowerLetter': return String.fromCharCode(96 + ((n - 1) % 26) + 1)
    case 'upperLetter': return String.fromCharCode(64 + ((n - 1) % 26) + 1)
    case 'lowerRoman': return toRoman(n, false)
    case 'upperRoman': return toRoman(n, true)
    default: return String(n)
  }
}

/** Expand a w:lvlText like "%1.%2" against the per-level counters. */
function expandLvlText(lvlText: string, counters: number[]): string {
  return lvlText.replace(/%(\d)/g, (_, d: string) => {
    const n = counters[Number(d) - 1]
    return n != null ? String(n) : ''
  })
}

/** Everything the renderer needs to lay out one paragraph, resolved from all layers. */
interface ResolvedPara {
  align?: ParaFormat['align']
  indentLeftPx?: number
  indentRightPx?: number
  indentFirstLinePx?: number
  lineSpacing?: number
  lineRule?: ParaFormat['lineRule']
  lineRawTwips?: number
  spaceBeforePx?: number
  spaceAfterPx?: number
  shadingFill?: string
  sizeHalfPoints?: number
  color?: string
  bold?: boolean
  italic?: boolean
  font?: string
}

/** Normalized paragraph format: both naming schemes flattened to twips fields. */
interface NormalFormat {
  align?: ParaFormat['align']
  indentLeftTwips?: number
  indentRightTwips?: number
  indentFirstLineTwips?: number
  lineSpacing?: number
  lineRule?: ParaFormat['lineRule']
  lineRawTwips?: number
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
  shadingFill?: string
  sizeHalfPoints?: number
  color?: string
  bold?: boolean
  italic?: boolean
  font?: string
}

type FormatLayer = Partial<ParaFormat> | Partial<ResolvedPara> | undefined

/**
 * ParaFormat (direct formatting) and StyleInfo.display (style chain) name the same
 * property differently (`spaceBefore` vs `spaceBeforeTwips`); read the fields off the
 * layer by shape so both sources merge through one code path.
 */
function normalizeFormat(layer: FormatLayer): NormalFormat {
  const n: NormalFormat = {}
  if (!layer) return n
  const v = layer as Record<string, unknown>
  const num = (k: string): number | undefined => (typeof v[k] === 'number' ? (v[k] as number) : undefined)
  const str = (k: string): string | undefined => (typeof v[k] === 'string' ? (v[k] as string) : undefined)
  const bool = (k: string): boolean | undefined => (typeof v[k] === 'boolean' ? (v[k] as boolean) : undefined)
  if (v.align != null) n.align = v.align as ParaFormat['align']
  n.indentLeftTwips = num('indentLeft') ?? num('indentLeftTwips')
  n.indentRightTwips = num('indentRight') ?? num('indentRightTwips')
  n.indentFirstLineTwips = num('indentFirstLine') ?? num('indentFirstLineTwips')
  n.lineSpacing = num('lineSpacing')
  n.lineRawTwips = num('lineRawTwips')
  if (v.lineRule != null) n.lineRule = v.lineRule as ParaFormat['lineRule']
  n.spaceBeforeTwips = num('spaceBefore') ?? num('spaceBeforeTwips')
  n.spaceAfterTwips = num('spaceAfter') ?? num('spaceAfterTwips')
  n.shadingFill = str('shadingFill')
  n.sizeHalfPoints = num('sizeHalfPoints')
  n.color = str('color')
  n.bold = bool('bold')
  n.italic = bool('italic')
  n.font = str('fontAscii') ?? str('font')
  return n
}

// ── tiny DOM helpers (standard API; no host prototype extensions) ──

function el<K extends keyof HTMLElementTagNameMap>(
  parent: Node,
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag)
  if (className) node.className = className
  parent.appendChild(node)
  return node
}

function div(parent: Node, className?: string): HTMLDivElement {
  return el(parent, 'div', className)
}

function span(parent: Node, className?: string): HTMLSpanElement {
  return el(parent, 'span', className)
}

/** mime of data URLs Chromium cannot decode natively (vector metafiles). */
const METAFILE_URL_RE = /^data:image\/(?:x-)?(?:emf|wmf|emz|wmz)[;,]/

/** Broken/external references arrive as ~empty dataUrls ("data:,"), not real image data. */
function isUsableImageDataUrl(url: string | undefined): url is string {
  return !!url && url.startsWith('data:image/') && url.length > 30
}

/** Decode a `data:` URL back to bytes + mime (metafile inputs are base64). */
function dataUrlBytes(url: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(url)
  if (!m) return null
  const bin = atob(m[2]!)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return { bytes, mime: m[1]! }
}

export class DocxRenderer {
  private doc: ParsedDoc
  private root: HTMLElement
  /** numbering counter state: `${abstractNumId}:${ilvl}` → count */
  private counters = new Map<string, number>()
  private defaults: ResolvedPara = { sizeHalfPoints: 22 }
  /** imgs whose source is an EMF/WMF data URL Chromium cannot decode; converted async after render */
  private metafileImgs: Array<{ img: HTMLImageElement; url: string }> = []

  constructor(doc: ParsedDoc, root: HTMLElement) {
    this.doc = doc
    this.root = root
  }

  render(): void {
    this.metafileImgs = []
    const paper = div(this.root, 'ovx-paper')
    this.applyDocDefaults()
    this.renderHeaderFooter(paper, 'header')
    const body = div(paper, 'ovx-body')
    for (const block of this.doc.blocks) {
      this.renderBlock(body, block)
    }
    this.renderHeaderFooter(paper, 'footer')
    void this.convertMetafiles()
  }

  /**
   * EMF/WMF pictures arrive as `data:image/emf…` URLs which Chromium cannot decode;
   * rasterize them with the engine's converter and swap the src in place. Failed
   * conversions degrade to the img's broken-image state.
   */
  private async convertMetafiles(): Promise<void> {
    const pending = this.metafileImgs
    this.metafileImgs = []
    await Promise.all(
      pending.map(async ({ img, url }) => {
        try {
          const parsed = dataUrlBytes(url)
          if (!parsed) return
          const png = await metafileToDataUrl(parsed.bytes, parsed.mime)
          if (png && img.isConnected) img.src = png
        } catch {
          // keep the original src; the browser shows its broken-image state
        }
      }),
    )
  }

  /** Create an <img> for a picture dataUrl, queueing metafiles for async conversion. */
  private makeImage(url: string): HTMLImageElement {
    const img = document.createElement('img')
    if (METAFILE_URL_RE.test(url)) {
      this.metafileImgs.push({ img, url })
    } else {
      img.src = url
    }
    return img
  }

  /** Approximate floating-image wrap with CSS float (left/right variants). */
  private applyImageWrap(el_: HTMLElement, wrap: Block['imageWrap'], distLeftEmu?: number, distRightEmu?: number): void {
    if (!wrap) return
    if (wrap.endsWith('-left')) el_.style.cssFloat = 'left'
    else if (wrap.endsWith('-right')) el_.style.cssFloat = 'right'
    else return
    // wp:wrap distances (EMU) become float margins so text keeps its gutter
    const PX_PER_EMU = 96 / 914400
    if (distLeftEmu) el_.style.marginLeft = `${Math.round(distLeftEmu * PX_PER_EMU)}px`
    if (distRightEmu) el_.style.marginRight = `${Math.round(distRightEmu * PX_PER_EMU)}px`
  }

  private applyDocDefaults(): void {
    const d: DocDefaults | undefined = this.doc.docDefaults
    const theme = this.doc.themeFonts
    const size = d?.sizeHalfPoints ?? 22
    this.defaults = {
      sizeHalfPoints: size,
      color: d?.color,
      font: d?.asciiFont ?? theme?.minor,
    }
    this.root.style.setProperty('--ovx-base-font', fontStack(this.defaults.font, theme?.eastAsia))
    this.root.style.setProperty('--ovx-base-size', pxOfHalfPoints(size))
    if (this.defaults.color) this.root.style.setProperty('--ovx-base-color', `#${this.defaults.color}`)
  }

  // ── header / footer strip ──

  private renderHeaderFooter(paper: HTMLElement, which: 'header' | 'footer'): void {
    const paras: HfParagraph[] | null | undefined =
      which === 'header' ? this.doc.headerParas : this.doc.footerParas
    const images = which === 'header' ? this.doc.headerImages : this.doc.footerImages
    const plainText = which === 'header' ? this.doc.headerText : this.doc.footerText
    if (!paras?.length && !images?.length && !plainText) return
    const strip = div(paper, `ovx-${which}`)
    if (paras?.length) {
      for (const para of paras) {
        const p = div(strip, 'ovx-hf-para')
        this.applyParaFormat(p, this.mergeParaFormats([this.defaults, para]))
        this.renderRuns(p, para.runs ?? [])
      }
    } else if (plainText) {
      div(strip).textContent = plainText.replace(/[\uE000\uE001]/g, '#')
    }
    for (const img of images ?? []) {
      if (!img.floating) strip.appendChild(this.makeImage(img.dataUrl))
    }
  }

  // ── blocks ──

  private renderBlock(body: HTMLElement, block: Block): void {
    if (block.hidden || block.invisibleMarker) return
    switch (block.type) {
      case 'heading':
        this.renderHeading(body, block)
        break
      case 'paragraph':
        this.renderParagraph(body, block)
        break
      case 'listItem':
        this.renderListItem(body, block)
        break
      case 'table':
        if (block.table) this.renderTable(body, block.table)
        break
      case 'image':
        this.renderImage(body, block)
        break
      case 'passthrough':
        this.renderPassthrough(body, block)
        break
    }
    for (const box of block.textboxes ?? []) this.renderTextbox(body, box)
  }

  private mergeParaFormats(layers: FormatLayer[]): ResolvedPara {
    const out: ResolvedPara = {}
    for (const layer of layers) {
      const n = normalizeFormat(layer)
      if (n.align != null) out.align = n.align
      if (n.indentLeftTwips != null) out.indentLeftPx = n.indentLeftTwips * PX_PER_TWIP
      if (n.indentRightTwips != null) out.indentRightPx = n.indentRightTwips * PX_PER_TWIP
      if (n.indentFirstLineTwips != null) out.indentFirstLinePx = n.indentFirstLineTwips * PX_PER_TWIP
      if (n.lineSpacing != null) out.lineSpacing = n.lineSpacing
      if (n.lineRule != null) out.lineRule = n.lineRule
      if (n.lineRawTwips != null) out.lineRawTwips = n.lineRawTwips
      if (n.spaceBeforeTwips != null) out.spaceBeforePx = n.spaceBeforeTwips * PX_PER_TWIP
      if (n.spaceAfterTwips != null) out.spaceAfterPx = n.spaceAfterTwips * PX_PER_TWIP
      if (n.shadingFill != null) out.shadingFill = n.shadingFill
      if (n.sizeHalfPoints != null) out.sizeHalfPoints = n.sizeHalfPoints
      if (n.color != null) out.color = n.color
      if (n.bold != null) out.bold = n.bold
      if (n.italic != null) out.italic = n.italic
      if (n.font != null) out.font = n.font
    }
    return out
  }

  private applyParaFormat(el_: HTMLElement, fmt: ResolvedPara): void {
    const s = el_.style
    if (fmt.align === 'center') s.textAlign = 'center'
    else if (fmt.align === 'right') s.textAlign = 'right'
    else if (fmt.align === 'justify' || fmt.align === 'distribute') s.textAlign = 'justify'
    if (fmt.indentLeftPx != null) s.marginLeft = `${fmt.indentLeftPx}px`
    if (fmt.indentRightPx != null) s.marginRight = `${fmt.indentRightPx}px`
    if (fmt.indentFirstLinePx != null) s.textIndent = `${fmt.indentFirstLinePx}px`
    if (fmt.lineRule === 'exact' && fmt.lineRawTwips != null) {
      s.lineHeight = `${fmt.lineRawTwips * PX_PER_TWIP}px`
    } else if (fmt.lineRule === 'atLeast' && fmt.lineRawTwips != null) {
      s.lineHeight = `max(${fmt.lineRawTwips * PX_PER_TWIP}px, normal)`
    } else if (fmt.lineSpacing != null) {
      s.lineHeight = String(fmt.lineSpacing)
    }
    if (fmt.spaceBeforePx != null) s.marginTop = `${fmt.spaceBeforePx}px`
    if (fmt.spaceAfterPx != null) s.marginBottom = `${fmt.spaceAfterPx}px`
    if (fmt.shadingFill) s.backgroundColor = `#${fmt.shadingFill}`
    if (fmt.sizeHalfPoints != null) s.fontSize = pxOfHalfPoints(fmt.sizeHalfPoints)
    if (fmt.color) s.color = `#${fmt.color}`
    if (fmt.font) s.fontFamily = fontStack(fmt.font)
  }

  private styleOf(styleId?: string): StyleInfo | undefined {
    return styleId ? this.doc.styles.get(styleId) : undefined
  }

  private renderHeading(body: HTMLElement, block: Block): void {
    const level = Math.min(Math.max(block.level ?? 1, 1), 6)
    const display = this.styleOf(block.styleId)?.display
    const resolved = this.mergeParaFormats([this.defaults, display, block.format])
    const p = div(body, `ovx-para ovx-h ovx-h${level}`)
    this.applyParaFormat(p, resolved)
    p.style.fontWeight = display?.bold ?? resolved.bold ?? true ? 'bold' : 'normal'
    if (!p.style.fontSize) p.style.fontSize = headingFallbackSize(level)
    this.renderRuns(p, block.runs ?? [])
  }

  private renderParagraph(body: HTMLElement, block: Block): void {
    const resolved = this.mergeParaFormats([this.defaults, this.styleOf(block.styleId)?.display, block.format])
    const p = div(body, 'ovx-para')
    this.applyParaFormat(p, resolved)
    if (block.format?.borders) applyParaBorders(p, block.format)
    this.renderRuns(p, block.runs ?? [])
    if (!p.childNodes.length) {
      // keep the Word line height of empty paragraphs
      p.textContent = '\u00a0'
      if (block.format?.emptyRunSizeHalfPoints) {
        p.style.fontSize = pxOfHalfPoints(block.format.emptyRunSizeHalfPoints)
      }
    }
  }

  private renderListItem(body: HTMLElement, block: Block): void {
    const list = block.list
    const resolved = this.mergeParaFormats([this.defaults, this.styleOf(block.styleId)?.display, block.format])
    const numDef = list ? this.doc.numbering.get(list.numId) : undefined
    const levelDef = list && numDef ? numDef.levels[list.ilvl] : undefined

    let markerText = '•'
    if (list && numDef) {
      const abstract = numDef.abstractNumId
      // deeper levels restart when an enclosing level increments
      for (const key of [...this.counters.keys()]) {
        if (key.startsWith(`${abstract}:`)) {
          const ilvl = Number(key.split(':')[1])
          if (ilvl > list.ilvl) this.counters.delete(key)
        }
      }
      const ckey = `${abstract}:${list.ilvl}`
      const start = numDef.startOverrides[list.ilvl] ?? levelDef?.start ?? 1
      const next = (this.counters.get(ckey) ?? start - 1) + 1
      this.counters.set(ckey, next)
      if (levelDef && levelDef.numFmt !== 'bullet') {
        const counters: number[] = []
        for (let l = 0; l <= list.ilvl; l++) {
          counters.push(this.counters.get(`${abstract}:${l}`) ?? start)
        }
        markerText = expandLvlText(levelDef.lvlText, counters)
        // non-decimal single-level formats (%1 with numFmt i/a/…) shape the whole marker
        if (levelDef.numFmt !== 'decimal') {
          markerText = markerText.replace(/\b\d+\b/, (m) => formatCounter(levelDef.numFmt, Number(m)))
        }
      } else if (levelDef?.lvlText) {
        markerText = levelDef.lvlText
      }
    }

    const p = div(body, 'ovx-para ovx-li')
    const hangingPx = Math.max(levelDef?.hanging != null ? levelDef.hanging * PX_PER_TWIP : 18, 12)
    p.style.marginLeft = `${Math.max(resolved.indentLeftPx ?? 0, hangingPx)}px`
    const marker = span(p, 'ovx-li-marker')
    marker.textContent = markerText
    marker.style.minWidth = `${hangingPx}px`
    if (levelDef?.szHalfPoints) marker.style.fontSize = pxOfHalfPoints(levelDef.szHalfPoints)
    this.applyParaFormat(p, { ...resolved, indentLeftPx: undefined, indentFirstLinePx: undefined })
    this.renderRuns(p, block.runs ?? [])
    if (!p.childNodes.length) p.textContent = '\u00a0'
  }

  private renderImage(body: HTMLElement, block: Block): void {
    if (block.brokenImage || !isUsableImageDataUrl(block.imageDataUrl)) {
      const chip = div(body, 'ovx-broken-image')
      chip.textContent = `[image unavailable${block.previewText ? `: ${block.previewText}` : ''}]`
      return
    }
    const wrap = div(body, 'ovx-image-block')
    this.applyImageWrap(wrap, block.imageWrap, block.imageWrapDistLeftEmu, block.imageWrapDistRightEmu)
    const floated = wrap.style.cssFloat !== ''
    if (!floated) {
      if (block.imageAlign === 'center') wrap.style.textAlign = 'center'
      else if (block.imageAlign === 'right') wrap.style.textAlign = 'right'
    }
    const img = this.makeImage(block.imageDataUrl)
    wrap.appendChild(img)
    if (block.imageWidthPx) img.style.width = `${block.imageWidthPx}px`
    if (block.imageHeightPx) img.style.height = `${block.imageHeightPx}px`
    img.style.maxWidth = '100%'
    if (block.imageBorder) {
      img.style.border = `${block.imageBorder.widthPt}px solid #${block.imageBorder.color}`
    }
    if (block.imageLeadingText) wrap.prepend(document.createTextNode(block.imageLeadingText))
  }

  private renderTextbox(body: HTMLElement, box: NonNullable<Block['textboxes']>[number]): void {
    const boxEl = div(body, 'ovx-textbox')
    if (box.widthPx) boxEl.style.width = `${box.widthPx}px`
    if (box.heightPx) boxEl.style.minHeight = `${box.heightPx}px`
    if (box.fill) boxEl.style.backgroundColor = `#${box.fill}`
    if (box.borderColor) boxEl.style.borderColor = `#${box.borderColor}`
    if (box.textColor) boxEl.style.color = `#${box.textColor}`
    for (const para of box.paras) {
      const p = div(boxEl, 'ovx-para')
      this.applyParaFormat(p, this.mergeParaFormats([this.defaults, this.styleOf(para.styleId)?.display, para]))
      if (box.vAlign === 'center') p.style.display = 'table-cell'
      this.renderRuns(p, para.runs)
    }
  }

  private renderPassthrough(body: HTMLElement, block: Block): void {
    const fd = block.fieldDisplay
    if (fd?.kind === 'tocLine') {
      const line = div(body, 'ovx-para ovx-toc-line')
      if (fd.level) line.style.paddingLeft = `${(fd.level - 1) * 18}px`
      line.style.display = 'flex'
      line.style.alignItems = 'baseline'
      line.style.gap = '4px'
      if (fd.num) span(line).textContent = `${fd.num} `
      span(line).textContent = fd.left ?? ''
      const dots = span(line, 'ovx-toc-dots')
      dots.style.flex = '1'
      dots.style.borderBottom = '1px dotted currentColor'
      dots.style.minHeight = '1em'
      if (fd.right) span(line).textContent = fd.right
      if (fd.szHalfPoints) line.style.fontSize = pxOfHalfPoints(fd.szHalfPoints)
      return
    }
    if (fd?.kind === 'text') {
      const p = div(body, 'ovx-para')
      if (fd.runs?.length) {
        for (const r of fd.runs) {
          const s = span(p)
          s.textContent = r.text
          if (r.szHalfPoints) s.style.fontSize = pxOfHalfPoints(r.szHalfPoints)
        }
      } else {
        p.textContent = fd.left ?? ''
      }
      if (fd.align === 'center') p.style.textAlign = 'center'
      else if (fd.align === 'right') p.style.textAlign = 'right'
      return
    }
    if (fd?.kind === 'pageBreak') {
      div(body, 'ovx-page-break')
      return
    }
    if (block.decorative) {
      const rule = div(body, 'ovx-rule')
      if (block.ruleColorHex) rule.style.borderColor = `#${block.ruleColorHex}`
      if (block.ruleThicknessPx) rule.style.borderTopWidth = `${block.ruleThicknessPx}px`
      return
    }
    const chip = div(body, 'ovx-chip')
    chip.textContent = block.label ?? 'Embedded object'
    if (block.previewText) {
      const preview = span(chip, 'ovx-chip-preview')
      preview.textContent = ` — ${block.previewText}`
    }
  }

  // ── tables ──

  private renderTable(body: HTMLElement, model: TableModel): void {
    const table = el(body, 'table', 'ovx-table')
    table.style.tableLayout = model.autoLayout ? 'auto' : 'fixed'
    if (model.widthPct != null) {
      table.style.width = `${Math.min(model.widthPct / 50, 100)}%`
    } else {
      table.style.width = '100%'
      const widths = model.colWidthsTwips ?? model.colWidthsPct
      if (widths?.length) {
        const total = widths.reduce((a, b) => a + b, 0) || 1
        const colgroup = el(table, 'colgroup')
        for (const w of widths) {
          const col = el(colgroup, 'col')
          col.style.width = `${((w / total) * 100).toFixed(2)}%`
        }
      }
    }
    if (model.fill) table.style.backgroundColor = `#${model.fill}`
    if (model.align === 'center') table.style.marginInline = 'auto'
    else if (model.align === 'right') table.style.marginInline = '0 0 auto auto'
    else if (model.indentTwips) table.style.marginLeft = `${model.indentTwips * PX_PER_TWIP}px`

    const rows = model.rows
    const rowCount = rows.length
    // columns claimed by an active vertical merge: col → (lastRow, colspan)
    const mergeUntil = new Map<number, { row: number; span: number }>()
    for (let r = 0; r < rowCount; r++) {
      const tr = el(table, 'tr')
      let c = 0
      for (const cell of rows[r] ?? []) {
        while (true) {
          const active = mergeUntil.get(c)
          if (active && active.row >= r) c += active.span
          else break
        }
        const colspan = cell.colSpan ?? 1
        if (cell.gridGap || cell.vMerge === 'continue') {
          c += colspan
          continue
        }
        const td = el(tr, 'td', 'ovx-td')
        if (colspan > 1) td.colSpan = colspan
        if (cell.vMerge === 'restart') {
          const spanRows = countVMergeRun(rows, r, c, colspan)
          if (spanRows > 1) {
            td.rowSpan = spanRows
            mergeUntil.set(c, { row: r + spanRows - 1, span: colspan })
          }
        }
        this.renderCell(td, model, cell, r, rowCount)
        c += colspan
      }
    }
  }

  private renderCell(
    td: HTMLTableCellElement,
    model: TableModel,
    cell: TableModel['rows'][number][number],
    rowIndex: number,
    rowCount: number,
  ): void {
    const style = this.styleOf(model.tblStyleId || undefined)?.tableDisplay
    const look = model.tableLook
    const fill = cell.fill
      ?? conditionalFill(style, look, rowIndex, rowCount)
      ?? style?.fill
    if (fill) td.style.backgroundColor = `#${fill}`
    if (cell.vAlign === 'center') td.style.verticalAlign = 'middle'
    else if (cell.vAlign === 'bottom') td.style.verticalAlign = 'bottom'
    if (cell.textDirection === 'tbRl') td.style.writingMode = 'vertical-rl'
    else if (cell.textDirection === 'btLr') td.style.writingMode = 'vertical-lr'

    let bold = cell.bold
    if (style && look?.firstRow && rowIndex === 0 && style.firstRow?.bold) bold = true
    const paras = cell.richParas ?? []
    if (!paras.length) {
      for (const text of cell.paras) {
        div(td, 'ovx-para').textContent = text
      }
    }
    for (const para of paras) {
      const p = div(td, 'ovx-para')
      const resolved = this.mergeParaFormats([this.defaults, this.styleOf(para.styleId)?.display, para])
      this.applyParaFormat(p, resolved)
      if (bold) p.style.fontWeight = 'bold'
      this.renderRuns(p, para.runs, bold ? { bold: true } : {})
    }
    for (const nested of cell.nestedTables ?? []) this.renderTable(td, nested)
    if (!td.childNodes.length) td.textContent = '\u00a0'
  }

  // ── runs ──

  private renderRuns(host: HTMLElement, runs: Run[], hints: { bold?: boolean } = {}): void {
    for (const run of runs) {
      if (run.vanish) continue
      if (run.del) continue // tracked deletions stay hidden in the preview
      if (run.image) {
        if (!isUsableImageDataUrl(run.image.dataUrl)) continue
        const img = this.makeImage(run.image.dataUrl)
        host.appendChild(img)
        this.applyImageWrap(img, run.image.wrap, run.image.wrapDistLeftEmu, run.image.wrapDistRightEmu)
        if (run.image.widthPx) img.style.width = `${run.image.widthPx}px`
        if (run.image.heightPx) img.style.height = `${run.image.heightPx}px`
        img.style.maxWidth = '100%'
        img.style.verticalAlign = 'text-bottom'
        continue
      }
      if (run.math) {
        span(host, 'ovx-math').textContent = run.text || '[formula]'
        continue
      }
      if (run.noteRef) {
        el(host, 'sup', 'ovx-note-ref').textContent = run.text || '*'
        continue
      }
      const display = this.styleOf(run.styleId)?.display
      const rawText = run.text
      if (!rawText) continue
      const text = run.caps === 'all' ? rawText.toUpperCase() : rawText
      if (run.link?.href) {
        const a = el(host, 'a', 'ovx-link')
        a.href = run.link.href
        a.target = '_blank'
        a.rel = 'noopener'
        a.textContent = text
        this.applyRunStyle(a, run, display, hints)
        continue
      }
      const s = span(host, 'ovx-run')
      s.textContent = text
      this.applyRunStyle(s, run, display, hints)
    }
  }

  private applyRunStyle(
    el_: HTMLElement,
    run: Run,
    display: StyleInfo['display'],
    hints: { bold?: boolean },
  ): void {
    const s = el_.style
    if (run.bold ?? display?.bold ?? hints.bold) s.fontWeight = 'bold'
    if (run.italic ?? display?.italic) s.fontStyle = 'italic'
    const underline = run.underline ?? display?.underline
    const strike = run.strike ?? display?.strike
    if (underline || strike) {
      s.textDecoration = [underline ? 'underline' : '', strike ? 'line-through' : '']
        .filter(Boolean)
        .join(' ')
    }
    if (run.ins) s.color = '#1a7f37'
    else {
      const color = run.color ?? display?.color
      if (color) s.color = `#${color}`
    }
    const size = run.sizeHalfPoints ?? display?.sizeHalfPoints
    if (size) s.fontSize = pxOfHalfPoints(size)
    const font = run.fontAscii ?? run.font ?? display?.fontAscii ?? display?.font
    if (font) s.fontFamily = fontStack(font)
    if (run.charSpacingTwips) s.letterSpacing = `${run.charSpacingTwips * PX_PER_TWIP}px`
    if (run.vertAlign === 'superscript') {
      s.verticalAlign = 'super'
      s.fontSize = '0.7em'
    } else if (run.vertAlign === 'subscript') {
      s.verticalAlign = 'sub'
      s.fontSize = '0.7em'
    }
    const highlightHex = run.highlight ? HIGHLIGHT[run.highlight] : undefined
    if (highlightHex) s.backgroundColor = `#${highlightHex}`
    else if (run.shading) s.backgroundColor = `#${run.shading}`
  }
}

/** Conditional fill from the table style, honoring w:tblLook switches. */
function conditionalFill(
  style: TableStyleDisplay | undefined,
  look: TableLook | undefined,
  rowIndex: number,
  rowCount: number,
): string | undefined {
  if (!style) return undefined
  if (look?.firstRow && rowIndex === 0) return style.firstRow?.fill
  if (look?.lastRow && rowIndex === rowCount - 1) return style.lastRow?.fill
  if (look?.bandedRows) return rowIndex % 2 === 1 ? style.band1Fill : style.band2Fill
  return undefined
}

/** Count how many following rows a vMerge 'restart' cell spans. */
function countVMergeRun(rows: TableModel['rows'], startRow: number, startCol: number, colspan: number): number {
  let span = 1
  for (let r = startRow + 1; r < rows.length; r++) {
    let c = 0
    let covered = false
    for (const cell of rows[r] ?? []) {
      const w = cell.colSpan ?? 1
      if (c + w > startCol && c < startCol + colspan) {
        covered = cell.vMerge === 'continue'
        break
      }
      c += w
    }
    if (!covered) break
    span++
  }
  return span
}

function applyParaBorders(p: HTMLElement, fmt: ParaFormat): void {
  const sides = fmt.borders ?? ''
  const lines = fmt.borderLines
  const widthPt = (fmt.borderStyle?.szEighths ?? 4) / 8
  const color = fmt.borderStyle?.color
  const map: Array<[string, 't' | 'b' | 'l' | 'r']> = [
    ['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r'],
  ]
  for (const [cssSide, key] of map) {
    if (!sides.includes(key)) continue
    const line = lines?.[key]
    p.style[`border${cap(cssSide)}` as 'borderTop'] =
      `${line?.szPt ?? widthPt}px solid #${line?.color ?? color ?? 'cccccc'}`
    p.style.paddingTop = p.style.paddingBottom = '2px'
  }
}

function cap(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1)
}

function headingFallbackSize(level: number): string {
  const pt = { 1: 32, 2: 26, 3: 22, 4: 20, 5: 18, 6: 16 }[level] ?? 20
  return `${(pt * 96) / 72}px`
}

function pxOfHalfPoints(hp: number): string {
  return `${(hp * PX_PER_HALF_POINT).toFixed(1)}px`
}

function fontStack(ascii?: string, eastAsia?: string): string {
  const parts = [ascii, eastAsia].filter(Boolean) as string[]
  parts.push('system-ui', 'sans-serif')
  return parts.map((f) => (/[\s",]/.test(f) ? `"${f}"` : f)).join(', ')
}
