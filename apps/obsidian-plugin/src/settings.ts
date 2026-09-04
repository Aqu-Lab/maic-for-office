import { App, PluginSettingTab, Setting } from 'obsidian'
import type OfficeViewerPlugin from './main'

export interface OfficeViewerSettings {
  /** Default pptx zoom: fit the slide to the view width when it opens */
  pptxFitWidth: boolean
  /** xlsx row/col count cap for the virtual table (guards against pathological sheets) */
  xlsxMaxRows: number
  xlsxMaxCols: number
}

export const DEFAULT_SETTINGS: OfficeViewerSettings = {
  pptxFitWidth: true,
  xlsxMaxRows: 5000,
  xlsxMaxCols: 200,
}

export async function loadSettings(plugin: OfficeViewerPlugin): Promise<OfficeViewerSettings> {
  const data = (await plugin.loadData()) as Partial<OfficeViewerSettings> | null
  return { ...DEFAULT_SETTINGS, ...data }
}

export async function saveSettings(plugin: OfficeViewerPlugin, settings: OfficeViewerSettings): Promise<void> {
  await plugin.saveData(settings)
}

export class OfficeViewerSettingTab extends PluginSettingTab {
  plugin: OfficeViewerPlugin

  constructor(app: App, plugin: OfficeViewerPlugin) {
    super(app, plugin)
    this.plugin = plugin
  }

  display(): void {
    const { containerEl } = this
    containerEl.empty()

    new Setting(containerEl).setName('Office Viewer').setHeading()

    new Setting(containerEl)
      .setName('Fit slides to view width')
      .setDesc('Open .pptx slides scaled to the pane width instead of 100%.')
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.pptxFitWidth).onChange(async (value) => {
          this.plugin.settings.pptxFitWidth = value
          await saveSettings(this.plugin, this.plugin.settings)
        }),
      )

    new Setting(containerEl)
      .setName('Spreadsheet row limit')
      .setDesc('Cap the rows rendered for one worksheet (performance guard).')
      .addText((text) =>
        text
          .setValue(String(this.plugin.settings.xlsxMaxRows))
          .onChange(async (value) => {
            const n = Number(value)
            if (Number.isFinite(n) && n > 0) {
              this.plugin.settings.xlsxMaxRows = Math.floor(n)
              await saveSettings(this.plugin, this.plugin.settings)
            }
          }),
      )
  }
}
