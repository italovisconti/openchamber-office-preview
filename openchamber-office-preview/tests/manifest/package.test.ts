import { describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { parseManifestJson } from '@openchamber/sdk/schemas';

describe('installable OpenChamber package manifest', () => {
  test('the extension manifest is accepted by the pinned SDK schema without a local build', async () => {
    const manifest = await readFile(resolve(import.meta.dir, '../../package.json'), 'utf8');
    const parsed = parseManifestJson(manifest);

    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error(`${parsed.code}: ${parsed.message}`);
    expect(parsed.version).toBe('0.1.0');
    expect(parsed.manifest.contributes.panel.id).toBe('office-preview');
    expect(parsed.manifest.contributes.fileEditors).toHaveLength(3);
  });

  test('the repository root Git-install manifest matches the package and points to built assets', async () => {
    const root = JSON.parse(await readFile(resolve(import.meta.dir, '../../../package.json'), 'utf8'));
    const nested = JSON.parse(await readFile(resolve(import.meta.dir, '../../package.json'), 'utf8'));
    expect(root.name).toBe(nested.name);
    expect(root.version).toBe(nested.version);
    expect(root.openchamber.contributes.panel).toEqual(nested.openchamber.contributes.panel);
    expect(root.openchamber.contributes.fileEditors).toHaveLength(3);
    for (const editor of root.openchamber.contributes.fileEditors) {
      const source = nested.openchamber.contributes.fileEditors.find((item: { id: string }) => item.id === editor.id);
      expect({ ...editor, entry: source.entry }).toEqual(source);
      expect(editor.entry).toBe(`openchamber-office-preview/${source.entry}`);
      expect(await Bun.file(resolve(import.meta.dir, '../../../', editor.entry)).exists()).toBe(true);
    }
  });
});
