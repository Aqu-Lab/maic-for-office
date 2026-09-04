import { Notice, TFile, Vault } from 'obsidian'

/** Read a vault file as bytes for the engine packages (they take Uint8Array). */
export async function readVaultBinary(vault: Vault, file: TFile): Promise<Uint8Array> {
  const buf = await vault.adapter.readBinary(file.path)
  return new Uint8Array(buf)
}

/** Show a user-facing error with the file name for context. */
export function showLoadError(file: TFile | null, err: unknown): void {
  const name = file ? `"${file.name}"` : 'file'
  const message = err instanceof Error ? err.message : String(err)
  console.error('[office-viewer]', err)
  new Notice(`Office Viewer: failed to open ${name} — ${message}`, 8000)
}
