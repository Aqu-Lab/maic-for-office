import { TFile, WorkspaceLeaf } from 'obsidian'
import { BaseOfficeView } from './base-view'
import type { XlsxDoc } from './render/xlsx-table'
import type OfficeViewerPlugin from '../main'

export const VIEW_TYPE_XLSX = 'office-viewer-xlsx'

const A1_COL = (c: number): string => {
  let s = ''
  c += 1
  while (c > 0) {
    const m = (c - 1) % 26
    s = String.fromCharCode(65 + m) + s
    c = Math.floor((c - 1) / 26)
  }
  return s
}

/**
 * Read-only .xlsx viewer: sheet tabs + sticky-header grid. Values and cached
 * formula results only; the Rust-sidecar fidelity of the Sheets app (styles,
 * charts, pivots) is out of scope here.
 */
export class XlsxView extends BaseOfficeView {
  private doc: XlsxDoc | null = null
  private tabHandler: ((ev: Event) => void) | null = null

  constructor(leaf: WorkspaceLeaf, plugin: OfficeViewerPlugin) {
    super(leaf, plugin)
  }

  getViewType(): string {
    return VIEW_TYPE_XLSX
  }

  getDisplayText(): string {
    const file = this.getFile()
    return file ? file.name : 'Spreadsheet'
  }

  getIcon(): string {
    return 'table'
  }

  protected extension(): string {
    return 'xlsx'
  }

  protected onCleanup(): void {
    this.doc = null
    this.tabHandler = null
  }

  protected async renderInto(container: HTMLElement, bytes: Uint8Array, _file: TFile): Promise<void> {
    const { parseXlsx } = await import('./render/xlsx-table')
    const doc = parseXlsx(bytes, this.plugin.settings.xlsxMaxRows, this.plugin.settings.xlsxMaxCols)
    this.doc = doc
    container.empty()
    const root = container.createDiv('ovx-xlsx')

    const tabs = root.createDiv('ovx-xlsx-tabs')
    const scroll = root.createDiv('ovx-xlsx-scroll')
    const meta = root.createDiv('ovx-xlsx-meta')

    const showSheet = (index: number) => {
      for (const tab of Array.from(tabs.children)) tab.removeClass('ovx-active')
      tabs.children[index]?.addClass('ovx-active')
      this.renderSheet(scroll, meta, doc.sheets[index]!)
    }
    doc.sheets.forEach((sheet, i) => {
      const tab = tabs.createDiv('ovx-xlsx-tab')
      tab.textContent = sheet.name
      tab.addEventListener('click', () => showSheet(i))
    })
    this.tabHandler = null
    showSheet(0)
  }

  private renderSheet(scroll: HTMLElement, meta: HTMLElement, sheet: XlsxDoc['sheets'][number]): void {
    scroll.empty()
    const table = scroll.createEl('table')
    const thead = table.createEl('thead')
    const headerRow = thead.createEl('tr')
    headerRow.createEl('th')
    for (let c = 0; c < (sheet.rows[0]?.length ?? 0); c++) {
      headerRow.createEl('th', { text: A1_COL(c) })
    }
    const tbody = table.createEl('tbody')
    sheet.rows.forEach((row, r) => {
      const tr = tbody.createEl('tr')
      tr.createEl('th', { text: String(r + 1) })
      for (let c = 0; c < row.length; c++) {
        const v = row[c] ?? ''
        const td = tr.createEl('td')
        td.textContent = v
        if (v !== '' && Number.isFinite(Number(v))) td.addClass('ovx-num')
      }
    })
    const hiddenRows = sheet.totalRows - sheet.rows.length
    const hiddenCols = sheet.totalCols - (sheet.rows[0]?.length ?? 0)
    const notes: string[] = [`${sheet.totalRows} rows × ${sheet.totalCols} cols`]
    if (hiddenRows > 0) notes.push(`${hiddenRows} rows not shown (limit in settings)`)
    if (hiddenCols > 0) notes.push(`${hiddenCols} cols not shown`)
    meta.setText(notes.join(' · '))
  }
}
