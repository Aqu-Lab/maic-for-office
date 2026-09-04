import { TFile, WorkspaceLeaf } from 'obsidian'
import { BaseOfficeView } from './base-view'
import { DocxRenderer } from './render/docx-renderer'
import type OfficeViewerPlugin from '../main'

export const VIEW_TYPE_DOCX = 'office-viewer-docx'

/** Read-only .docx preview backed by the docx-engine block tree. */
export class DocxView extends BaseOfficeView {
  constructor(leaf: WorkspaceLeaf, plugin: OfficeViewerPlugin) {
    super(leaf, plugin)
  }

  getViewType(): string {
    return VIEW_TYPE_DOCX
  }

  getDisplayText(): string {
    const file = this.getFile()
    return file ? file.name : 'Word document'
  }

  getIcon(): string {
    return 'file-text'
  }

  protected extension(): string {
    return 'docx'
  }

  protected async renderInto(container: HTMLElement, bytes: Uint8Array, _file: TFile): Promise<void> {
    const { parseDocx } = await import('@maic/docx-engine')
    const doc = await parseDocx(bytes)
    container.empty()
    const root = container.createDiv('ovx-docx')
    new DocxRenderer(doc, root).render()
  }
}
