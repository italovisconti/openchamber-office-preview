import { Unzip, UnzipInflate, UnzipPassThrough, zipSync } from 'fflate';

const MAX_ARCHIVE = 20_000_000;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_ENTRY = 16 * 1024 * 1024;
const MAX_ENTRIES = 2_000;
// fflate may emit a large output burst for one push. Keep each compressed input
// slice small so its transient inflate allocation is bounded as well.
const INPUT_CHUNK = 4 * 1024;
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX_XLSX_ROWS = 100_000;
const MAX_XLSX_COLUMNS = 500;
const MAX_XLSX_CELLS = 250_000;
const MAX_XLSX_MERGE_AREA = 10_000;
const MAX_XLSX_MERGE_COUNT = 10_000;

/** Conservative OOXML gate. Only ordinary ZIP/deflate OOXML parts are accepted;
 * no extraction to disk or attempt to sanitize active package features is made. */
export async function prepareOfficeArchive(
  bytes: Uint8Array,
  format: 'xlsx' | 'docx' | 'pptx',
  options: { signal?: AbortSignal } = {},
): Promise<Uint8Array> {
  if (bytes.byteLength > MAX_ARCHIVE) throw new Error('Office archive exceeds compressed size limit');
  if (options.signal?.aborted) throw abortError();
  const centralNames = inspectCentralDirectory(bytes);
  const parts = new Map<string, Uint8Array>();
  const originalNames = new Map<string, string>();
  const seen = new Set<string>();
  let total = 0;
  let count = 0;

  await new Promise<void>(async (resolve, reject) => {
    let stopped = false;
    const fail = (error: Error) => {
      if (stopped) return;
      stopped = true;
      reject(error);
    };
    const unzip = new Unzip((file) => {
      if (stopped) return;
      if (options.signal?.aborted) { fail(abortError()); return; }
      count++;
      const name = normalizeName(file.name);
      if (count > MAX_ENTRIES) { fail(new Error('Too many ZIP entries')); return; }
      if (seen.has(name)) { fail(new Error('Duplicate normalized ZIP path')); return; }
      seen.add(name);
      const chunks: Uint8Array[] = [];
      let size = 0;
      file.ondata = (error, chunk, final) => {
        if (stopped) return;
        if (error) { fail(new Error('Invalid or truncated ZIP entry')); return; }
        if (options.signal?.aborted) { fail(abortError()); return; }
        size += chunk.byteLength;
        total += chunk.byteLength;
        if (size > MAX_ENTRY || total > MAX_TOTAL) { fail(new Error('Office archive expansion limit exceeded')); return; }
        chunks.push(chunk);
        if (final) {
          const data = new Uint8Array(size);
          let offset = 0;
          for (const part of chunks) { data.set(part, offset); offset += part.byteLength; }
            parts.set(name, data);
            originalNames.set(name, file.name);
        }
      };
      file.start();
    });
    unzip.register(UnzipPassThrough);
    unzip.register(UnzipInflate);
    try {
      for (let offset = 0, chunkCount = 0; offset < bytes.length; offset += INPUT_CHUNK, chunkCount++) {
        if (stopped) return;
        if (options.signal?.aborted) throw abortError();
        const end = Math.min(offset + INPUT_CHUNK, bytes.length);
        unzip.push(bytes.subarray(offset, end), end >= bytes.length);
        if (stopped) return;
        if (chunkCount % 16 === 15) await new Promise<void>((yieldControl) => setTimeout(yieldControl, 0));
      }
      if (stopped) return;
      if (seen.size !== centralNames.size || [...centralNames].some((name) => !seen.has(name))) throw new Error('Truncated or inconsistent ZIP central directory');
      if (parts.size !== centralNames.size || [...centralNames].some((name) => !parts.has(name))) throw new Error('Incomplete ZIP entry data');
      resolve();
    } catch (error) { fail(error instanceof Error ? error : new Error(String(error))); }
  });

  const has = (key: string) => parts.has(key.toLowerCase());
  const required = format === 'xlsx' ? 'xl/workbook.xml' : format === 'docx' ? 'word/document.xml' : 'ppt/presentation.xml';
  if (!has('[Content_Types].xml') || !has('_rels/.rels') || !has(required)) {
    throw new Error('OOXML package does not match the requested format');
  }
  const contentTypes = text(parts.get('[content_types].xml')!);
  const expectedType = format === 'xlsx' ? /spreadsheetml\.sheet\.main\+xml/i : format === 'docx' ? /wordprocessingml\.document\.main\+xml/i : /presentationml\.presentation\.main\+xml/i;
  if (!expectedType.test(contentTypes)) throw new Error('Incompatible OOXML main content type');
  if (has('word/document.xml') !== (format === 'docx') || has('xl/workbook.xml') !== (format === 'xlsx') || has('ppt/presentation.xml') !== (format === 'pptx')) {
    throw new Error('Ambiguous OOXML package contains another document type');
  }
  const chartWorkbooks = format === 'pptx' ? validatedChartWorkbooks(parts) : new Set<string>();
  if (format === 'xlsx') validateXlsxBudgets(parts);
  for (const [name, data] of parts) {
    const lower = name.toLowerCase();
    if (chartWorkbooks.has(name) || (name === 'ppt/embeddings' && chartWorkbooks.size > 0)) continue;
    if (/\.(?:bin|vba|xll|svg|fntdata)$/i.test(lower) || /(?:^|\/)(?:activex|embeddings|activeX|embeddings)(?:\/|$)/i.test(name) || /(?:^|\/)customui\//i.test(name)) {
      throw new Error('Active or unverifiable OOXML content is not supported');
    }
    if (lower.endsWith('.rels') || lower.endsWith('.xml')) {
      const xml = text(data);
      if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('XML DTD/entities are not supported');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.querySelector('parsererror')) throw new Error('Malformed OOXML XML');
      if (doc.querySelector('Relationship[TargetMode="External" i]')) throw new Error('External OOXML relationships are not supported');
      const root = doc.documentElement;
      const activeElements = ['altChunk', 'oleObject', 'control', 'legacyDrawing', 'svgBlip', 'vbaProject', 'externalLink'];
      if (activeElements.some((local) => doc.getElementsByTagNameNS('*', local).length > 0)) throw new Error('Active or unverifiable OOXML feature is not supported');
      if (lower.endsWith('.rels') && [...doc.getElementsByTagNameNS('*', 'Relationship')].some((rel) => {
        const type = rel.getAttribute('Type') ?? '';
        if (/\/package$/i.test(type) && chartWorkbooks.has(resolveSafeChartTarget(name, rel.getAttribute('Target') ?? ''))) return false;
        return /\/(?:oleobject|control|externallink|attachedtemplate|package|image\.svg\+xml)$/i.test(type);
      })) throw new Error('Dangerous OOXML relationship');
      if (root.localName === 'Types' && [...doc.getElementsByTagNameNS('*', 'Default')].some((entry) => /vba|activex/i.test(entry.getAttribute('ContentType') ?? ''))) throw new Error('Active or unverifiable OOXML content is not supported');
      if (lower.endsWith('.rels') && /(?:Target\s*=\s*["'][^"']*\.(?:html?|js|svg|bin|vba)|TargetMode\s*=\s*["']External)/i.test(xml)) throw new Error('Dangerous OOXML relationship');
    }
  }
  const canonical: Record<string, Uint8Array> = {};
  for (const [name, data] of parts) canonical[originalNames.get(name)!] = data;
  return zipSync(canonical, { level: 6 });
}

/** Compatibility wrapper for callers interested only in validation. */
export async function validateOfficeArchive(bytes: Uint8Array, format: 'xlsx' | 'docx' | 'pptx', options: { signal?: AbortSignal } = {}): Promise<void> {
  await prepareOfficeArchive(bytes, format, options);
}

function validateXlsxBudgets(parts: Map<string, Uint8Array>): void {
  let cells = 0, merges = 0, mergeArea = 0;
  let worksheets = 0;
  for (const [name, bytes] of parts) {
    // ExcelJS recognizes this substring even with a prefix/suffix. Reject aliases
    // so no worksheet can bypass the anchored budget checks below.
    if (/xl\/worksheets\/sheet\d+\.xml/i.test(name)
      && !/^xl\/worksheets\/sheet\d+\.xml$/i.test(name)) {
      throw new Error('Non-canonical XLSX worksheet path');
    }
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(name)) continue;
    if (++worksheets > 200) throw new Error('XLSX has too many worksheets');
    const doc = parseXml(bytes);
    const dimension = [...doc.getElementsByTagNameNS('*', 'dimension')][0]?.getAttribute('ref');
    if (dimension) checkRange(dimension, MAX_XLSX_ROWS * MAX_XLSX_COLUMNS);
    const rows = [...doc.getElementsByTagNameNS('*', 'row')];
    if (rows.length > MAX_XLSX_ROWS) throw new Error('XLSX row limit exceeded');
    for (const row of rows) {
      const index = Number(row.getAttribute('r'));
      if (!Number.isInteger(index) || index < 1 || index > MAX_XLSX_ROWS) throw new Error('XLSX row range limit exceeded');
    }
    for (const cell of [...doc.getElementsByTagNameNS('*', 'c')]) {
      const address = cell.getAttribute('r');
      if (!address) throw new Error('XLSX cell is missing an address');
      checkRange(address, 1);
      if (++cells > MAX_XLSX_CELLS) throw new Error('XLSX cell budget exceeded');
    }
    const sheetMerges = [...doc.getElementsByTagNameNS('*', 'mergeCell')];
    merges += sheetMerges.length;
    if (merges > MAX_XLSX_MERGE_COUNT) throw new Error('XLSX merge count limit exceeded');
    for (const merge of sheetMerges) {
      const area = checkRange(merge.getAttribute('ref') ?? '', MAX_XLSX_MERGE_AREA);
      mergeArea += area;
      if (mergeArea > MAX_XLSX_MERGE_AREA) throw new Error('XLSX merged-cell materialization budget exceeded');
    }
  }
}
function checkRange(ref: string, maxArea: number): number {
  const pieces = ref.split(':');
  if (pieces.length > 2) throw new Error('Invalid XLSX cell range');
  const start = parseCellAddress(pieces[0]);
  const end = parseCellAddress(pieces[1] ?? pieces[0]);
  if (end.row < start.row || end.col < start.col || end.row > MAX_XLSX_ROWS || end.col > MAX_XLSX_COLUMNS) throw new Error('XLSX row or column range limit exceeded');
  const area = (end.row - start.row + 1) * (end.col - start.col + 1);
  if (area > maxArea) throw new Error('XLSX range materialization budget exceeded');
  return area;
}
function parseCellAddress(address: string): { row: number; col: number } {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(address);
  if (!match) throw new Error('Invalid XLSX cell address');
  let col = 0;
  for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (!Number.isSafeInteger(row)) throw new Error('XLSX row range limit exceeded');
  return { row, col };
}

/** Whitelist only a chart's cached-data workbook; the renderer strips this part before loading. */
function validatedChartWorkbooks(parts: Map<string, Uint8Array>): Set<string> {
  const accepted = new Set<string>();
  for (const [relsPath, bytes] of parts) {
    const match = /^ppt\/charts\/_rels\/chart(\d+)\.xml\.rels$/i.exec(relsPath);
    if (!match) continue;
    const chartPath = `ppt/charts/chart${match[1]}.xml`;
    const chartBytes = parts.get(chartPath.toLowerCase());
    if (!chartBytes) continue;
    const relDoc = parseXml(bytes), chartDoc = parseXml(chartBytes);
    for (const rel of [...relDoc.getElementsByTagNameNS(PACKAGE_REL, 'Relationship')]) {
      if (rel.getAttribute('Type') !== `${OFFICE_REL}/package` || rel.hasAttribute('TargetMode')) continue;
      const target = resolveSafeChartTarget(relsPath, rel.getAttribute('Target') ?? '');
      if (!/^ppt\/embeddings\/microsoft_excel_worksheet\d+\.xlsx$/i.test(target)) continue;
      const id = rel.getAttribute('Id');
      const externalData = [...chartDoc.getElementsByTagNameNS('*', 'externalData')].some((el) => el.getAttributeNS(OFFICE_REL, 'id') === id);
      if (id && externalData && parts.has(target.toLowerCase())) accepted.add(target.toLowerCase());
    }
  }
  return accepted;
}
function resolveSafeChartTarget(source: string, target: string): string {
  if (!target || target.startsWith('/') || /^[a-z]+:/i.test(target) || target.includes('\\')) return '';
  const sourcePart = source.replace('/_rels/', '/').replace(/\.rels$/i, '');
  const parts = sourcePart.split('/'); parts.pop();
  for (const piece of target.split('/')) {
    if (!piece || piece === '.') continue;
    if (piece === '..') { if (!parts.length) return ''; parts.pop(); }
    else parts.push(piece);
  }
  return parts.join('/').toLowerCase();
}
function parseXml(data: Uint8Array): Document {
  const xml = text(data);
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('XML DTD/entities are not supported');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.querySelector('parsererror')) throw new Error('Malformed OOXML XML');
  return doc;
}

function normalizeName(name: string): string {
  if (!name || name.includes('\\') || name.startsWith('/') || /^[a-z]:/i.test(name) || name.includes('\0')) throw new Error('Unsafe ZIP path');
  const pieces = name.split('/');
  if (pieces.some((part) => part === '.' || part === '..')) throw new Error('Unsafe ZIP path');
  const normalized = pieces.filter(Boolean).join('/').normalize('NFC').toLowerCase();
  if (!normalized) throw new Error('Unsafe ZIP path');
  return normalized;
}
function text(bytes: Uint8Array): string { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
function abortError(): Error { return new DOMException('Validation aborted', 'AbortError'); }

function inspectCentralDirectory(bytes: Uint8Array): Set<string> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (p: number) => view.getUint16(p, true);
  const u32 = (p: number) => view.getUint32(p, true);
  let eocd = -1;
  for (let p = bytes.length - 22, min = Math.max(0, bytes.length - 65_557); p >= min; p--) if (u32(p) === 0x06054b50) { eocd = p; break; }
  if (eocd < 0 || eocd + 22 + u16(eocd + 20) !== bytes.length) throw new Error('Truncated ZIP or missing end-of-central-directory');
  const disk = u16(eocd + 4), centralDisk = u16(eocd + 6), diskCount = u16(eocd + 8), count = u16(eocd + 10), centralSize = u32(eocd + 12), centralOffset = u32(eocd + 16);
  if (disk || centralDisk || diskCount !== count || count === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) throw new Error('Multi-disk or ZIP64 archives are not supported');
  if (count > MAX_ENTRIES || centralOffset + centralSize !== eocd) throw new Error('Invalid ZIP central directory');
  const names = new Set<string>();
  let p = centralOffset;
  for (let i = 0; i < count; i++) {
    if (p + 46 > eocd || u32(p) !== 0x02014b50) throw new Error('Truncated ZIP central directory');
    const flags = u16(p + 8), method = u16(p + 10), compressed = u32(p + 20), expanded = u32(p + 24), n = u16(p + 28), x = u16(p + 30), c = u16(p + 32);
    if (flags & 1) throw new Error('Encrypted ZIP entries are not supported');
    if (method !== 0 && method !== 8) throw new Error('Unsupported ZIP compression method');
    if (compressed === 0xffffffff || expanded === 0xffffffff || p + 46 + n + x + c > eocd) throw new Error('ZIP64 or truncated ZIP entry');
    const name = normalizeName(new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(p + 46, p + 46 + n)));
    if (names.has(name)) throw new Error('Duplicate normalized ZIP path');
    names.add(name);
    p += 46 + n + x + c;
  }
  if (p !== eocd) throw new Error('Invalid ZIP central directory length');
  return names;
}
