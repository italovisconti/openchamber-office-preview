import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, strToU8 } from 'fflate';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { prepareOfficeArchive, validateOfficeArchive } from '../../src/security/preflight.js';

// Requires a browser-like DOMParser (e.g. jsdom supplied by the test harness).
function packageZip(format: 'xlsx' | 'docx' | 'pptx', replacements: Record<string, string> = {}) {
  const main = format === 'xlsx' ? ['xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'] : format === 'docx' ? ['word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'] : ['ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'];
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/${main[0]}" ContentType="${main[1]}"/></Types>`),
    '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
    [main[0]]: strToU8('<root/>'),
  };
  for (const [name, value] of Object.entries(replacements)) files[name] = strToU8(value);
  return zipSync(files);
}
function makeAliasedPayload(xml: string): Uint8Array {
  const data = strToU8(xml), name = strToU8('word/document.xml');
  const record = new Uint8Array(30 + name.length + data.length);
  const view = new DataView(record.buffer);
  view.setUint32(0, 0x04034b50, true); view.setUint16(4, 20, true);
  view.setUint32(18, data.length, true); view.setUint32(22, data.length, true);
  view.setUint16(26, name.length, true); record.set(name, 30); record.set(data, 30 + name.length);
  const result = new Uint8Array(8 + record.length); result.set(record, 8); return result;
}
function findZipRecord(bytes: Uint8Array, signature: number, filename: string): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset + 30 <= bytes.length; offset++) {
    if (view.getUint32(offset, true) !== signature) continue;
    const nameLength = view.getUint16(offset + (signature === 0x04034b50 ? 26 : 28), true);
    const nameOffset = offset + (signature === 0x04034b50 ? 30 : 46);
    if (new TextDecoder().decode(bytes.subarray(nameOffset, nameOffset + nameLength)) === filename) return offset;
  }
  throw new Error(`ZIP record not found: ${filename}`);
}

for (const format of ['xlsx', 'docx', 'pptx'] as const) test(`accepts minimal ${format}`, async () => validateOfficeArchive(packageZip(format), format));
test('rejects wrong format with the expected contract error', async () => assert.rejects(validateOfficeArchive(packageZip('docx'), 'xlsx'), /OOXML package does not match the requested format/));
test('rejects external relationships with a user-facing diagnostic', async () => assert.rejects(validateOfficeArchive(packageZip('docx', { '_rels/.rels': '<Relationships><Relationship Target="https://example.test" TargetMode="External"/></Relationships>' }), 'docx'), /referencia a contenido externo/));
test('rejects active package parts with a user-facing diagnostic', async () => assert.rejects(validateOfficeArchive(packageZip('pptx', { 'ppt/embeddings/oleObject1.bin': 'x' }), 'pptx'), /contenido activo o incrustado/));
test('canonical output contains only locally inspected parts and preserves original part names', async () => {
  const source = packageZip('docx', { 'word/document.xml': '<w:document xmlns:w="urn:w"/>', 'Word/Custom.XML': '<custom/>' });
  const canonical = await prepareOfficeArchive(source, 'docx');
  const { unzipSync, strFromU8 } = await import('fflate');
  const entries = unzipSync(canonical);
  assert.ok(entries['word/document.xml']);
  assert.equal(strFromU8(entries['word/document.xml']), '<w:document xmlns:w="urn:w"/>');
  assert.equal(strFromU8(entries['Word/Custom.XML']), '<custom/>');
});
test('canonicalization prevents central-directory offsets from aliasing unchecked padding as a part', async () => {
  const malicious = '<w:document xmlns:w="urn:w"><w:altChunk/></w:document>';
  const prefix = '<w:document xmlns:w="urn:w">', suffix = '</w:document>';
  const safe = prefix + ' '.repeat(malicious.length - prefix.length - suffix.length) + suffix;
  const source = zipSync({
    '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
    'word/document.xml': strToU8(safe),
    'padding.dat': makeAliasedPayload(malicious),
  }, { level: 0 });
  const forged = source.slice();
  const docCentral = findZipRecord(forged, 0x02014b50, 'word/document.xml');
  const padLocal = findZipRecord(forged, 0x04034b50, 'padding.dat');
  const dv = new DataView(forged.buffer, forged.byteOffset, forged.byteLength);
  const padData = padLocal + 30 + dv.getUint16(padLocal + 26, true) + dv.getUint16(padLocal + 28, true);
  dv.setUint32(docCentral + 42, padData + 8, true);
  const maliciousSeenByCentralReader = await (await JSZip.loadAsync(forged)).file('word/document.xml')!.async('text');
  assert.match(maliciousSeenByCentralReader, /altChunk/);
  const canonical = await prepareOfficeArchive(forged, 'docx');
  const canonicalText = await (await JSZip.loadAsync(canonical)).file('word/document.xml')!.async('text');
  assert.doesNotMatch(canonicalText, /altChunk/);
});
test('permits only a chart-linked embedded XLSX workbook for cached chart data', async () => {
  const chartNs = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const chartXml = `<c:chartSpace xmlns:c="${chartNs}" xmlns:r="${officeRel}"><c:chart><c:plotArea/><c:externalData r:id="rIdWorkbook"/></c:chart></c:chartSpace>`;
  const chartRels = `<Relationships xmlns="${relNs}"><Relationship Id="rIdWorkbook" Type="${officeRel}/package" Target="../embeddings/Quarterly Summary.xlsx"/></Relationships>`;
  const additions = {
    'ppt/presentation.xml': '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="1" r:id="rSlide"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${relNs}"><Relationship Id="rSlide" Target="slides/slide1.xml"/></Relationships>`,
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
    'ppt/charts/chart1.xml': chartXml,
    'ppt/charts/_rels/chart1.xml.rels': chartRels,
    'ppt/embeddings/Quarterly Summary.xlsx': 'fake nested workbook bytes',
  };
  const source = packageZip('pptx', additions);
  await validateOfficeArchive(source, 'pptx');
  const { visibleSlidesCopy } = await import('../../src/viewers/pptx-package.js');
  const rendered = await JSZip.loadAsync(await visibleSlidesCopy(source));
  assert.equal(rendered.file('ppt/embeddings/Quarterly Summary.xlsx'), null);
  assert.doesNotMatch(await rendered.file('ppt/charts/_rels/chart1.xml.rels')!.async('text'), /<Relationship\b/);
  assert.doesNotMatch(await rendered.file('ppt/charts/chart1.xml')!.async('text'), /externalData/);
});
test('accepts a legitimate chart workbook with an arbitrary embedded XLSX basename', async () => {
  const chartNs = 'http://schemas.openxmlformats.org/drawingml/2006/chart', officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const chartXml = `<c:chartSpace xmlns:c="${chartNs}" xmlns:r="${officeRel}"><c:chart><c:externalData r:id="rChartData"/></c:chart></c:chartSpace>`;
  const chartRels = `<Relationships xmlns="${relNs}"><Relationship Id="rChartData" Type="${officeRel}/package" Target="../embeddings/Monthly Results.xlsx"/></Relationships>`;
  const source = packageZip('pptx', {
    'ppt/charts/chart1.xml': chartXml,
    'ppt/charts/_rels/chart1.xml.rels': chartRels,
    'ppt/embeddings/Monthly Results.xlsx': 'not expanded by preflight',
  });
  await validateOfficeArchive(source, 'pptx');
});
test('a separate non-chart package relationship cannot reuse the chart workbook allowance', async () => {
  const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const chart = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRel}"><c:externalData r:id="chartData"/></c:chartSpace>`;
  const chartRels = `<Relationships xmlns="${relNs}"><Relationship Id="chartData" Type="${officeRel}/package" Target="../embeddings/Workbook.xlsx"/></Relationships>`;
  const slideRels = `<Relationships xmlns="${relNs}"><Relationship Id="otherUse" Type="${officeRel}/package" Target="../embeddings/Workbook.xlsx"/></Relationships>`;
  await assert.rejects(validateOfficeArchive(packageZip('pptx', {
    'ppt/charts/chart1.xml': chart,
    'ppt/charts/_rels/chart1.xml.rels': chartRels,
    'ppt/slides/_rels/slide1.xml.rels': slideRels,
    'ppt/embeddings/Workbook.xlsx': 'nested archive payload',
}), 'pptx'), /relación a un archivo incrustado/);
});
test('removes only typed presentation printer settings from the canonical renderer archive', async () => {
  const packageNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const printerType = 'application/vnd.openxmlformats-officedocument.presentationml.printerSettings';
  const source = packageZip('pptx', {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/printerSettings/printerSettings3.bin" ContentType="${printerType}"/></Types>`,
    'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${packageNs}"><Relationship Id="rPrinter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="printerSettings/printerSettings3.bin"/></Relationships>`,
    'ppt/printerSettings/printerSettings3.bin': 'opaque settings bytes',
  });
  const canonical = await prepareOfficeArchive(source, 'pptx');
  const original = await JSZip.loadAsync(source), rendered = await JSZip.loadAsync(canonical);
  expect(original.file('ppt/printerSettings/printerSettings3.bin')).toBeTruthy();
  expect(rendered.file('ppt/printerSettings/printerSettings3.bin')).toBeNull();
  expect(await rendered.file('ppt/_rels/presentation.xml.rels')!.async('text')).not.toMatch(/<Relationship\b/);
  expect(await rendered.file('[Content_Types].xml')!.async('text')).not.toContain('printerSettings3.bin');
});
test('does not allow arbitrary binary parts through printer-settings metadata', async () => {
  const packageNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const printerType = 'application/vnd.openxmlformats-officedocument.presentationml.printerSettings';
  const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/embeddings/notPrinter.bin" ContentType="${printerType}"/></Types>`;
  const rels = `<Relationships xmlns="${packageNs}"><Relationship Id="rPrinter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="embeddings/notPrinter.bin"/></Relationships>`;
  await assert.rejects(validateOfficeArchive(packageZip('pptx', { '[Content_Types].xml': contentTypes, 'ppt/_rels/presentation.xml.rels': rels, 'ppt/embeddings/notPrinter.bin': 'binary' }), 'pptx'), /no compatible|contenido activo o incrustado/);
  const wrongMime = contentTypes.replace(printerType, 'application/octet-stream').replace('ppt/embeddings/notPrinter.bin', 'ppt/printerSettings/printerSettings3.bin');
  const validLocation = rels.replace('embeddings/notPrinter.bin', 'printerSettings/printerSettings3.bin');
  await assert.rejects(validateOfficeArchive(packageZip('pptx', { '[Content_Types].xml': wrongMime, 'ppt/_rels/presentation.xml.rels': validLocation, 'ppt/printerSettings/printerSettings3.bin': 'binary' }), 'pptx'), /no compatible|contenido activo o incrustado/);
});
test('rejects an embedded workbook not referenced by a chart and arbitrary package relationships', async () => {
  const base = packageZip('pptx', { 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx': 'nested workbook' });
  await assert.rejects(validateOfficeArchive(base, 'pptx'), /contenido activo o incrustado/);
  const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Excel_Worksheet1.xlsx"/></Relationships>';
  await assert.rejects(validateOfficeArchive(packageZip('pptx', { 'ppt/charts/chart1.xml': '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>', 'ppt/charts/_rels/chart1.xml.rels': rels, 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx': 'nested workbook' }), 'pptx'), /relación a un archivo incrustado/);
});
test('rejects package relationships without a chart externalData, outside chart rels, or with traversal/invalid targets', async () => {
  const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const cases = [
    { relPath: 'ppt/charts/_rels/chart1.xml.rels', target: '../../embeddings/Workbook.xlsx', chart: `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRel}"/>` },
    { relPath: 'ppt/charts/_rels/chart1.xml.rels', target: '../embeddings/Workbook.zip', chart: `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRel}"><c:externalData r:id="r1"/></c:chartSpace>` },
    { relPath: 'ppt/charts/_rels/chart1.xml.rels', target: '../embeddings/Workbook.xlsx', chart: `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRel}"/>` },
    { relPath: 'ppt/slides/_rels/slide1.xml.rels', target: '../embeddings/Workbook.xlsx', chart: '' },
  ];
  for (const item of cases) {
    const relationship = `<Relationships xmlns="${relNs}"><Relationship Id="r1" Type="${officeRel}/package" Target="${item.target}"/></Relationships>`;
    await assert.rejects(validateOfficeArchive(packageZip('pptx', {
      'ppt/charts/chart1.xml': item.chart,
      [item.relPath]: relationship,
      'ppt/embeddings/Workbook.xlsx': 'nested archive payload',
    }), 'pptx'));
  }
  const external = `<Relationships xmlns="${relNs}"><Relationship Id="r1" Type="${officeRel}/package" Target="https://example.test/book.xlsx" TargetMode="External"/></Relationships>`;
  await assert.rejects(validateOfficeArchive(packageZip('pptx', { 'ppt/charts/chart1.xml': `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRel}"><c:externalData r:id="r1"/></c:chartSpace>`, 'ppt/charts/_rels/chart1.xml.rels': external }), 'pptx'), /referencia a contenido externo/);
  const ole = `<Relationships xmlns="${relNs}"><Relationship Id="r1" Type="${officeRel}/oleObject" Target="../embeddings/object1.bin"/></Relationships>`;
  await assert.rejects(validateOfficeArchive(packageZip('pptx', { 'ppt/slides/_rels/slide1.xml.rels': ole, 'ppt/embeddings/object1.bin': 'ole' }), 'pptx'), /relación a un archivo incrustado/);
});
test('enforces actual inflated byte budget', async () => assert.rejects(validateOfficeArchive(zipSync({ 'padding.txt': new Uint8Array(16 * 1024 * 1024 + 1) }), 'docx'), /expansion limit/));
test('honors cancellation', async () => { const controller = new AbortController(); controller.abort(); await assert.rejects(validateOfficeArchive(packageZip('docx'), 'docx', { signal: controller.signal }), { name: 'AbortError' }); });
test('yields between bounded ZIP input batches so cancellation can be observed', async () => {
  const controller = new AbortController();
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(validateOfficeArchive(zipSync({ 'padding.bin': new Uint8Array(512 * 1024) }, { level: 0 }), 'docx', { signal: controller.signal }), { name: 'AbortError' });
});
test('allows ordinary DOCX typography elements', async () => validateOfficeArchive(packageZip('docx', { 'word/styles.xml': '<w:styles xmlns:w="urn:w"><w:style><w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Arial"/></w:rPr></w:style></w:styles>' }), 'docx'));
test('allows standard package metadata relationships', async () => validateOfficeArchive(packageZip('xlsx', { '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' }), 'xlsx'));
test('allows an unused SVG MIME declaration, not SVG package content', async () => {
  const types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  await validateOfficeArchive(packageZip('docx', { '[Content_Types].xml': types }), 'docx');
  await assert.rejects(validateOfficeArchive(packageZip('docx', { '[Content_Types].xml': types, 'word/media/image.svg': '<svg/>' }), 'docx'), /contenido activo o incrustado/);
});
test('rejects XLSX huge sparse positions and merges before renderer parsing', async () => {
  const sheet = (body: string) => `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</worksheet>`;
  await assert.rejects(validateOfficeArchive(packageZip('xlsx', { 'xl/worksheets/sheet1.xml': sheet('<sheetData><row r="100001"><c r="A100001"/></row></sheetData>') }), 'xlsx'), /XLSX row range limit/);
  await assert.rejects(validateOfficeArchive(packageZip('xlsx', { 'xl/worksheets/sheet1.xml': sheet('<mergeCells><mergeCell ref="A1:XFD1048576"/></mergeCells>') }), 'xlsx'), /XLSX row or column range limit/);
  await assert.rejects(validateOfficeArchive(packageZip('xlsx', { 'xl/worksheets/sheet1.xml': sheet('<sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells><mergeCell ref="A1:Z500"/></mergeCells>') }), 'xlsx'), /XLSX range materialization budget/);
});
test('rejects worksheet filename aliases recognized by ExcelJS before loading them', async () => {
  const book = new ExcelJS.Workbook();
  book.addWorksheet('Data').getCell('A1').value = 'safe';
  const source = await book.xlsx.writeBuffer();
  for (const alias of ['xl/worksheets/sheet1.xml.extra', 'prefix/xl/worksheets/sheet1.xml']) {
    const zip = await JSZip.loadAsync(source);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('text');
    zip.remove('xl/worksheets/sheet1.xml');
    zip.file(alias, xml.replace('</worksheet>', '<mergeCells><mergeCell ref="A1:B6000"/></mergeCells></worksheet>'));
    await assert.rejects(prepareOfficeArchive(await zip.generateAsync({ type: 'uint8array' }), 'xlsx'), /Non-canonical XLSX worksheet path/);
  }
});
test('rejects truncated archives with a ZIP diagnostic', async () => assert.rejects(validateOfficeArchive(packageZip('xlsx').subarray(0, -8), 'xlsx'), /Truncated ZIP|end-of-central-directory/));
test('rejects encrypted ZIP metadata', async () => {
  const zip = packageZip('docx');
  const copy = zip.slice();
  for (let i = 0; i < copy.length - 4; i++) {
    const sig = new DataView(copy.buffer).getUint32(i, true);
    if (sig === 0x04034b50) copy[i + 6] |= 1;
    if (sig === 0x02014b50) copy[i + 8] |= 1;
  }
  await assert.rejects(validateOfficeArchive(copy, 'docx'), /Encrypted ZIP entries/);
});
test('rejects ZIP64 sentinel metadata', async () => {
  const copy = packageZip('docx').slice();
  for (let i = 0; i < copy.length - 4; i++) if (new DataView(copy.buffer).getUint32(i, true) === 0x02014b50) new DataView(copy.buffer).setUint32(i + 20, 0xffffffff, true);
  await assert.rejects(validateOfficeArchive(copy, 'docx'), /ZIP64/);
});
test('rejects duplicate normalized names', async () => {
  const copy = zipSync({ 'a.txt': strToU8('one'), 'b.txt': strToU8('two') });
  const renamed = copy.slice();
  for (let i = 0; i < renamed.length - 4; i++) {
    const sig = new DataView(renamed.buffer).getUint32(i, true);
    if (sig === 0x04034b50 || sig === 0x02014b50) {
      const nameLength = new DataView(renamed.buffer).getUint16(i + (sig === 0x04034b50 ? 26 : 28), true);
      const start = i + (sig === 0x04034b50 ? 30 : 46);
      if (new TextDecoder().decode(renamed.subarray(start, start + nameLength)) === 'b.txt') renamed[start] = 97;
    }
  }
  await assert.rejects(validateOfficeArchive(renamed, 'docx'), /Duplicate normalized ZIP path/);
});

test('parses every entry before active-feature checks in a padded 35 KB ZIP', async () => {
  const padding = new Uint8Array(34 * 1024);
  for (let i = 0; i < padding.length; i++) padding[i] = (i * 73 + 19) & 0xff;
  const zip = zipSync({
    'padding.txt': padding,
    '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
    'word/document.xml': strToU8('<w:document xmlns:w="urn:word"><w:body><w:altChunk/></w:body></w:document>'),
  }, { level: 0 });
  assert.ok(zip.byteLength > 35_000 && zip.byteLength < 40_000, `fixture ZIP is ${zip.byteLength} bytes`);
  await assert.rejects(validateOfficeArchive(zip, 'docx'), /característica activa no compatible/);
});
