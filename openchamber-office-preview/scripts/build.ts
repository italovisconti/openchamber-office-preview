import { build } from 'esbuild';
import { mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { $ } from 'bun';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
await $`bun run ${resolve(root, 'scripts/generate-pptx-worker.ts')}`;
await mkdir(resolve(root, 'editor'), { recursive: true });

const entries = {
  xlsx: resolve(root, 'src/entries/xlsx.ts'),
  docx: resolve(root, 'src/entries/docx.ts'),
  pptx: resolve(root, 'src/entries/pptx.ts'),
};
for (const [name, entry] of Object.entries(entries)) {
  await build({
    entryPoints: [entry],
    outfile: resolve(root, `editor/${name}.js`),
    bundle: true,
    format: 'iife',
    platform: 'browser',
    target: 'es2022',
    legalComments: 'none',
    sourcemap: false,
    minify: true,
  });
}

console.log('Built three classic IIFE editors and local stylesheets.');
