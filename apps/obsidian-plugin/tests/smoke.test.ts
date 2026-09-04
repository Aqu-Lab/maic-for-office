/**
 * Smoke tests for the three viewer pipelines, run against the real engines:
 * docx parse → renderer DOM; pptx parse → render tree; xlsx parse → cell grid.
 * These run in jsdom; the Konva layer is not exercised here (no canvas).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { createBlankPptx, openPptx } from '@maic/pptx-engine'
import { buildBlankDocx, parseDocx } from '@maic/docx-engine'
import { buildRenderSlide, HeuristicMetrics } from '@maic/pptx-render'
import * as XLSX from 'xlsx'
import { DocxRenderer } from '../src/views/render/docx-renderer'
import { parseXlsx } from '../src/views/render/xlsx-table'

const REPO_ROOT = join(__dirname, '../../..')

describe('docx pipeline', () => {
  it('renders a fixture document into paper DOM', async () => {
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, 'fixtures/generated/simple.docx')))
    const doc = await parseDocx(bytes)
    const host = document.createElement('div')
    new DocxRenderer(doc, host).render()

    const paper = host.querySelector('.ovx-paper')
    expect(paper).toBeTruthy()
    // the fixture has visible content; every body block maps to a para/table/image/chip
    const content = host.querySelectorAll('.ovx-para, .ovx-table, .ovx-image-block, .ovx-chip, .ovx-h')
    expect(content.length).toBeGreaterThan(0)
  })

  it('renders a generated blank document', async () => {
    const blank = await buildBlankDocx()
    const doc = await parseDocx(blank)
    const host = document.createElement('div')
    new DocxRenderer(doc, host).render()
    expect(host.querySelector('.ovx-paper')).toBeTruthy()
  })
})

describe('pptx pipeline', () => {
  it('opens a generated deck and builds render trees', async () => {
    const bytes = await createBlankPptx()
    const opened = await openPptx(bytes)
    expect(opened.deck.slides.length).toBeGreaterThan(0)
    const metrics = new HeuristicMetrics()
    const tree = buildRenderSlide(opened.deck.slides[0]!, opened.deck.size, {
      fitWidthPx: 1280,
      metrics,
    })
    expect(tree.widthPx).toBeGreaterThan(0)
    expect(tree.heightPx).toBeGreaterThan(0)
  })
})

describe('xlsx pipeline', () => {
  it('parses a workbook into sheet grids', () => {
    const wb = XLSX.utils.book_new()
    const ws = XLSX.utils.aoa_to_sheet([
      ['Name', 'Score'],
      ['alpha', 90],
      ['beta', 85],
    ])
    XLSX.utils.book_append_sheet(wb, ws, 'Grades')
    const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }))
    const doc = parseXlsx(bytes, 5000, 200)
    expect(doc.names).toEqual(['Grades'])
    const sheet = doc.sheets[0]!
    expect(sheet.rows[0]).toEqual(['Name', 'Score'])
    expect(sheet.rows[1]?.[1]).toBe('90')
  })
})
