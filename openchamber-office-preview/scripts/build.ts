import { build } from 'esbuild';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
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
  if (name === 'docx') {
    // DOCX's CSP forbids linked stylesheets. Keep its inline CSS synchronized
    // with the freshly compiled UI instead of retaining an older theme.
    const htmlPath = resolve(root, 'editor/docx.html');
    const css = await readFile(resolve(root, 'editor/docx.css'), 'utf8');
    const html = await readFile(htmlPath, 'utf8');
    if (!/<style>[\s\S]*?<\/style>/.test(html)) throw new Error('DOCX inline stylesheet is missing');
    await writeFile(htmlPath, html.replace(/<style>[\s\S]*?<\/style>/, () => `<style>${css}</style>`));
  }
}

console.log('Built three classic IIFE editors and local stylesheets.');
