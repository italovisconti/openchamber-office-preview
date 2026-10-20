import { readFile, readdir, mkdir, rm, writeFile, copyFile } from 'node:fs/promises';
import { resolve, relative, join, dirname, sep } from 'node:path';
import { zipSync, type Zippable } from 'fflate';

const root = resolve(import.meta.dir, '..');
const staging = resolve(root, 'dist/extension');
const manifest = JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8')) as { name: string; version: string };
const output = resolve(root, `dist/${manifest.name}-${manifest.version}.zip`);
await rm(resolve(root, 'dist'), { recursive: true, force: true });
await mkdir(staging, { recursive: true });
const pkg = JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8'));
const installManifest = {
  name: pkg.name,
  version: pkg.version,
  type: pkg.type,
  license: pkg.license,
  engines: pkg.engines,
  dependencies: pkg.dependencies,
  openchamber: pkg.openchamber,
};
await writeFile(resolve(staging, 'package.json'), JSON.stringify(installManifest, null, 2) + '\n');
await copyFile(resolve(root, 'README.md'), resolve(staging, 'README.md'));
await copyFile(resolve(root, 'LICENSE'), resolve(staging, 'LICENSE'));
await copyFile(resolve(root, 'THIRD_PARTY_NOTICES.md'), resolve(staging, 'THIRD_PARTY_NOTICES.md'));
await cpTree(resolve(root, 'editor'), resolve(staging, 'editor'));

const licenses = resolve(staging, 'licenses');
await mkdir(licenses, { recursive: true });
const nodeModules = resolve(root, 'node_modules');
const packageFiles = await findLicenseFiles(nodeModules);
for (const path of packageFiles) {
  const rel = relative(nodeModules, path);
  const destinationRel = rel.split(sep).map((part) => part === 'node_modules' ? 'nested-packages' : part).join(sep);
  const destination = join(licenses, destinationRel);
  await mkdir(dirname(destination), { recursive: true });
  await copyFile(path, destination);
}
const workerNotice = resolve(nodeModules, '@file-viewer/pptx/dist/worker.THIRD_PARTY_LICENSES.txt');
await copyFile(workerNotice, resolve(licenses, 'file-viewer-pptx-worker-THIRD_PARTY_LICENSES.txt'));

const files = await walkFiles(staging);
if (files.length > 500) throw new Error(`Install archive has ${files.length} files (limit 500).`);
const zipInput: Zippable = {};
let unpacked = 0;
for (const path of files.sort()) {
  const name = relative(staging, path).split(sep).join('/');
  const bytes = new Uint8Array(await readFile(path));
  unpacked += bytes.byteLength;
  zipInput[name] = bytes;
}
if (unpacked > 40 * 1024 * 1024) throw new Error(`Install archive expands to ${unpacked} bytes (limit 40 MiB).`);
const zip = zipSync(zipInput, { level: 9, mtime: new Date('1980-01-01T00:00:00.000Z') });
if (zip.byteLength > 20_000_000) throw new Error(`Install ZIP is ${zip.byteLength} bytes (limit 20,000,000).`);
await writeFile(output, zip);
console.log(`Packaged ${files.length} files; ZIP ${zip.byteLength} bytes; unpacked ${unpacked} bytes: ${output}`);

async function cpTree(from: string, to: string): Promise<void> {
  await mkdir(to, { recursive: true });
  for (const entry of await readdir(from, { withFileTypes: true })) {
    const source = join(from, entry.name), destination = join(to, entry.name);
    if (entry.isDirectory()) await cpTree(source, destination);
    else await copyFile(source, destination);
  }
}
async function walkFiles(dir: string): Promise<string[]> {
  const files: string[] = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...await walkFiles(path));
    else if (entry.isFile()) files.push(path);
  }
  return files;
}
async function findLicenseFiles(dir: string): Promise<string[]> {
  const files: string[] = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    if (entry.name === '.bin') continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...await findLicenseFiles(path));
    else if (/^(?:LICENSE|LICENCE|NOTICE|COPYING)(?:\.|$)/i.test(entry.name)) files.push(path);
  }
  return files;
}
