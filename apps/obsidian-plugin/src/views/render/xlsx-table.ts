/**
 * xlsx read-only table — SheetJS (Community Edition) parse → virtualized DOM table.
 *
 * The upstream Sheets app parses .xlsx through its Rust sidecar, which cannot ship
 * inside an Obsidian plugin; SheetJS CE covers values, formulas (cached results),
 * dates, and merged ranges, which is the fidelity bar for viewing. Cell styling is
 * out of scope for the first version.
 */
import { read, utils, type WorkBook } from 'xlsx'

export interface XlsxDoc {
  /** sheet names in workbook order */
  names: string[]
  /** cells of each sheet as displayed strings, rows of columns (trimmed to the data range) */
  sheets: XlsxSheetData[]
  mergedCount: number
}

export interface XlsxSheetData {
  name: string
  rows: string[][]
  /** physical row/col counts the data range covers (for the footer note) */
  totalRows: number
  totalCols: number
}

export function parseXlsx(bytes: Uint8Array, maxRows: number, maxCols: number): XlsxDoc {
  const wb: WorkBook = read(bytes, { type: 'array', cellDates: true })
  const names = wb.SheetNames
  const sheets = names.map((name) => {
    const ws = wb.Sheets[name]
    if (!ws) return { name, rows: [], totalRows: 0, totalCols: 0 }
    const ref = ws['!ref'] ?? 'A1'
    const range = utils.decode_range(ref)
    const totalRows = range.e.r - range.s.r + 1
    const totalCols = range.e.c - range.s.c + 1
    // dense array of displayed values, capped for pathological sheets
    const rows = utils.sheet_to_json<string[]>(ws, {
      header: 1,
      raw: false,
      defval: '',
      blankrows: true,
    })
    const capped = rows.slice(0, maxRows).map((r) => r.slice(0, maxCols).map(cellText))
    return { name, rows: capped, totalRows, totalCols }
  })
  const mergedCount = names.reduce((acc, n) => acc + (wb.Sheets[n]?.['!merges']?.length ?? 0), 0)
  return { names, sheets, mergedCount }
}

function cellText(v: unknown): string {
  if (v == null) return ''
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  return String(v)
}
