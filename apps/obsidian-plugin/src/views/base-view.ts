import { FileView, TFile, WorkspaceLeaf } from 'obsidian'
import { readVaultBinary, showLoadError } from '../file-io'
import type OfficeViewerPlugin from '../main'

/**
 * Shared shell for the three office views: loading state, error card,
 * reload handling, and DOM teardown. Subclasses implement renderInto()
 * for their format.
 */
export abstract class BaseOfficeView extends FileView {
  protected plugin: OfficeViewerPlugin
  private contentRoot: HTMLElement | null = null
  private loadToken = 0

  constructor(leaf: WorkspaceLeaf, plugin: OfficeViewerPlugin) {
    super(leaf)
    this.plugin = plugin
    this.allowNoFile = true
  }

  protected abstract extension(): string
  protected abstract renderInto(container: HTMLElement, bytes: Uint8Array, file: TFile): Promise<void>
  protected onCleanup(): void {}

  getFile(): TFile | null {
    const file = this.file
    return file && file.extension.toLowerCase() === this.extension() ? file : null
  }

  async onOpen(): Promise<void> {
    const root = this.contentEl
    root.empty()
    root.addClass('office-viewer-root')
    this.contentRoot = root.createDiv('office-viewer-content')
    // Obsidian opens the view BEFORE the file is assigned (onLoadFile runs later),
    // so this usually renders nothing yet; the real render happens in onLoadFile.
    if (this.getFile()) await this.reload()
  }

  /** Obsidian assigns the file and calls this after onOpen — the actual render trigger. */
  async onLoadFile(file: TFile): Promise<void> {
    await super.onLoadFile(file)
    await this.reload()
  }

  async reload(): Promise<void> {
    const file = this.getFile()
    if (!this.contentRoot || !file) return
    const token = ++this.loadToken
    this.showLoading(file)
    try {
      const bytes = await readVaultBinary(this.app.vault, file)
      if (token !== this.loadToken) return
      const host = this.contentRoot
      host.empty()
      await this.renderInto(host, bytes, file)
    } catch (err) {
      if (token !== this.loadToken) return
      this.showError(file, err)
    }
  }

  async onClose(): Promise<void> {
    this.loadToken++
    this.onCleanup()
    this.contentRoot = null
    this.contentEl.empty()
  }

  private showLoading(file: TFile): void {
    if (!this.contentRoot) return
    this.contentRoot.empty()
    const box = this.contentRoot.createDiv('office-viewer-status')
    box.createSpan({ text: `Loading ${file.name}…` })
  }

  private showError(file: TFile, err: unknown): void {
    showLoadError(file, err)
    if (!this.contentRoot) return
    this.contentRoot.empty()
    const box = this.contentRoot.createDiv('office-viewer-status office-viewer-error')
    box.createSpan({
      text: `Could not render ${file.name}. ${(err instanceof Error ? err.message : String(err))}`,
    })
    const retry = box.createEl('button', { text: 'Retry' })
    retry.addEventListener('click', () => {
      void this.reload().catch(() => {})
    })
  }
}
