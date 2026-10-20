import { readFile, writeFile } from 'node:fs/promises';

const workerPath = new URL('../node_modules/@file-viewer/pptx/dist/worker/pptx.worker.js', import.meta.url);
const worker = await readFile(workerPath, 'utf8');
if (!worker.trim()) throw new Error(`PPTX worker is empty: ${workerPath.pathname}`);
const output = new URL('../src/pptx-worker-source.ts', import.meta.url);
await writeFile(output, `// Generated from @file-viewer/pptx@3.2.2; do not edit.\nexport const pptxWorkerSource = ${JSON.stringify(worker)};\n`);
console.log(`Embedded local PPTX worker source (${Buffer.byteLength(worker)} bytes).`);
