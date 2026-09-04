/**
 * Obsidian plugin bundle — mirrors the official obsidian-sample-plugin esbuild setup.
 *
 * Output: dist/main.js (+ manifest.json + styles.css copied alongside). When
 * OBSIDIAN_TEST_VAULT (env) or testVault (package.json "officeViewer" key) points at a
 * vault, the three files are also copied into <vault>/.obsidian/plugins/maic-for-office/
 * so a Community-plugins reload picks the build up immediately.
 *
 * Engine packages are consumed as workspace source (their exports point at src/index.ts),
 * so esbuild compiles everything into the single main.js Obsidian expects.
 */
import * as esbuild from 'esbuild'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const dist = path.join(here, 'dist')

const watch = process.argv.includes('--watch')

/** node: builtins resolve at runtime inside Obsidian's Electron renderer (desktop) */
const external = ['obsidian', 'electron', 'node:*']

async function copyToVault() {
  const envVault = process.env.OBSIDIAN_TEST_VAULT
  let vault = envVault ?? ''
  if (!vault) {
    try {
      const pkg = JSON.parse(fs.readFileSync(path.join(here, 'package.json'), 'utf8'))
      vault = pkg.officeViewer?.testVault ?? ''
    } catch {
      /* no package.json — skip */
    }
  }
  if (!vault) return null
  const target = path.join(vault, '.obsidian', 'plugins', 'maic-for-office')
  fs.mkdirSync(target, { recursive: true })
  for (const f of ['main.js', 'manifest.json', 'styles.css']) {
    fs.copyFileSync(path.join(dist, f), path.join(target, f))
  }
  return target
}

const ctx = await esbuild.context({
  entryPoints: [path.join(here, 'src', 'main.ts')],
  bundle: true,
  outfile: path.join(dist, 'main.js'),
  external,
  format: 'cjs',
  target: 'es2022',
  logLevel: 'info',
  sourcemap: watch ? 'inline' : false,
  treeShaking: true,
  // Obsidian desktop ships a current Chromium; the DOM + canvas APIs Konva needs are all there
  conditions: ['browser'],
  define: { 'process.env.NODE_ENV': '"production"' },
})

fs.mkdirSync(dist, { recursive: true })
fs.copyFileSync(path.join(here, 'manifest.json'), path.join(dist, 'manifest.json'))
fs.copyFileSync(path.join(here, 'styles.css'), path.join(dist, 'styles.css'))

if (watch) {
  await ctx.watch()
  const target = await copyToVault()
  console.log(`[office-viewer] watching; ${target ? `installed to ${target}` : 'set OBSIDIAN_TEST_VAULT to auto-install into a vault'}`)
} else {
  await ctx.rebuild()
  await ctx.dispose()
  const target = await copyToVault()
  console.log(`[office-viewer] built dist/main.js${target ? ` → installed to ${target}` : ''}`)
}
