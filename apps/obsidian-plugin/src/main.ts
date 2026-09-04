/**
 * Office Viewer — an Obsidian plugin that views .docx / .pptx / .xlsx attachments
 * with the GenOffice engine packages (parse in-process, read-only render).
 *
 * Each format gets its own ItemView type; registerExtensions maps the file
 * extension onto the view, so clicking an attachment in the vault opens it.
 */
import { Plugin, TFile } from 'obsidian'
import { OfficeViewerSettingTab, loadSettings, type OfficeViewerSettings } from './settings'
import { DocxView, VIEW_TYPE_DOCX } from './views/docx-view'
import { PptxView, VIEW_TYPE_PPTX } from './views/pptx-view'
import { XlsxView, VIEW_TYPE_XLSX } from './views/xlsx-view'

export default class OfficeViewerPlugin extends Plugin {
  /** Intentionally narrows the base class's `settings?: unknown` (TS2612). */
  declare settings: OfficeViewerSettings

  async onload(): Promise<void> {
    this.settings = await loadSettings(this)

    this.registerView(VIEW_TYPE_DOCX, (leaf) => new DocxView(leaf, this))
    this.registerView(VIEW_TYPE_PPTX, (leaf) => new PptxView(leaf, this))
    this.registerView(VIEW_TYPE_XLSX, (leaf) => new XlsxView(leaf, this))

    // registerExtensions makes Office Viewer the default opener for these files;
    // users can still route a file elsewhere via "Open in default app".
    this.registerExtensions(['docx'], VIEW_TYPE_DOCX)
    this.registerExtensions(['pptx'], VIEW_TYPE_PPTX)
    this.registerExtensions(['xlsx'], VIEW_TYPE_XLSX)

    this.addSettingTab(new OfficeViewerSettingTab(this.app, this))

    this.addCommand({
      id: 'reload-active-office-view',
      name: 'Reload active office view',
      checkCallback: (checking) => {
        const leaf = this.app.workspace.getActiveViewOfType(DocxView)
          ?? this.app.workspace.getActiveViewOfType(PptxView)
          ?? this.app.workspace.getActiveViewOfType(XlsxView)
        if (!leaf) return false
        if (!checking) void leaf.reload()
        return true
      },
    })
  }

  onunload(): void {
    // Obsidian detaches registered views' leaves itself; DOM cleanup happens in each view's onClose
  }
}

export type { OfficeViewerSettings }
export type OfficePluginFile = TFile
