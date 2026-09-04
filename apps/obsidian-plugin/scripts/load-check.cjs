/** Dev-only: load dist/main.js against a stubbed obsidian module to surface init-time crashes. */
const Module = require('module')
const orig = Module._load
Module._load = function (request) {
  if (request === 'obsidian') {
    return new Proxy({}, {
      get: (t, k) => {
        if (k === 'Plugin') {
          return class Plugin {
            addRibbonIcon() {}
            addCommand() {}
            registerView() {}
            registerExtensions() {}
            addSettingTab() {}
            registerEvent() {}
            loadData() { return Promise.resolve(null) }
            saveData() { return Promise.resolve() }
          }
        }
        if (k === 'ItemView' || k === 'FileView') {
          return class View {
            constructor() {
              this.contentEl = {
                addClass() {},
                createDiv() { return { createDiv() { return { createDiv() {}, style: {} } }, style: {} } },
                empty() {},
              }
            }
          }
        }
        if (k === 'PluginSettingTab') return class PluginSettingTab { constructor() {} }
        if (k === 'Setting') {
          return class Setting {
            setName() { return this }
            setDesc() { return this }
            addToggle() { return this }
            addText() { return this }
          }
        }
        if (k === 'Notice') return class Notice {}
        if (k === 'TFile' || k === 'Vault' || k === 'WorkspaceLeaf' || k === 'Menu' || k === 'Modal') return class Generic {}
        return class Fallback {}
      },
    })
  }
  return orig.apply(this, arguments)
}

const path = require('path')
try {
  const m = require(path.join(__dirname, '..', 'dist', 'main.js'))
  console.log('LOADED OK, exports:', Object.keys(m))
} catch (e) {
  console.error('LOAD FAILED:', e.message)
  console.error(e.stack.split('\n').slice(0, 6).join('\n'))
}
