import { describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { parseManifestJson } from '@openchamber/sdk/schemas';

describe('installable OpenChamber package manifest', () => {
  test('the exact package envelope is accepted by the pinned SDK schema', async () => {
    const manifest = await readFile(resolve(import.meta.dir, '../../dist/extension/package.json'), 'utf8');
    const parsed = parseManifestJson(manifest);

    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error(`${parsed.code}: ${parsed.message}`);
    expect(parsed.version).toBe('0.1.0');
    expect(parsed.manifest.contributes.panel.id).toBe('office-preview');
    expect(parsed.manifest.contributes.fileEditors).toHaveLength(3);
  });
});
