import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { visibleSlidesCopy } from '../../src/viewers/pptx-package';

const NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="1" r:id="r1"/><p:sldId id="2" r:id="r2"/><p:sldId id="3" r:id="r3"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="r1" Target="slides/slide1.xml"/><Relationship Id="r2" Target="slides/slide2.xml"/><Relationship Id="r3" Target="slides/slide3.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="' + NS + '"><p:cSld/></p:sld>');
  zip.file('ppt/slides/slide2.xml', '<p:sld xmlns:p="' + NS + '" show="0"><p:cSld/></p:sld>');
  zip.file('ppt/slides/slide3.xml', '<p:sld xmlns:p="' + NS + '"><p:cSld/></p:sld>');
  zip.file('ppt/slides/slide9.xml', '<p:sld xmlns:p="' + NS + '"><p:cSld/></p:sld>');
  return zip.generateAsync({ type: 'uint8array' });
}
describe('PPTX visible slide package copy', () => {
  test('keeps visible order and removes hidden and orphan slide parts without mutating source', async () => {
    const source = await fixture(), original = source.slice();
    const output = await visibleSlidesCopy(source);
    const result = await JSZip.loadAsync(output);
    const xml = await result.file('ppt/presentation.xml')!.async('text');
    expect([...xml.matchAll(/r:id="(r\d+)"/g)].map((m) => m[1])).toEqual(['r1', 'r3']);
    expect(result.file('ppt/slides/slide1.xml')).toBeTruthy();
    expect(result.file('ppt/slides/slide2.xml')).toBeNull();
    expect(result.file('ppt/slides/slide9.xml')).toBeNull();
    expect(await result.file('[Content_Types].xml')!.async('text')).not.toContain('slide2.xml');
    expect(await result.file('[Content_Types].xml')!.async('text')).not.toContain('slide9.xml');
    expect(source).toEqual(original);
  });

  test('strips chart-linked workbook with arbitrary case-preserved name and override', async () => {
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = await JSZip.loadAsync(await fixture());
    const types = await zip.file('[Content_Types].xml')!.async('text');
    zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/ppt/embeddings/Quarterly Summary.XLSX" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>'));
    zip.file('ppt/charts/chart1.xml', `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${rel}"><c:chart><c:externalData r:id="rChart"/></c:chart></c:chartSpace>`);
    zip.file('ppt/charts/_rels/chart1.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rChart" Type="${rel}/package" Target="../embeddings/Quarterly Summary.xlsx"/></Relationships>`);
    zip.file('ppt/embeddings/Quarterly Summary.XLSX', 'not opened as a nested archive');
    const output = await visibleSlidesCopy(await zip.generateAsync({ type: 'uint8array' }));
    const rendered = await JSZip.loadAsync(output);
    expect(rendered.file('ppt/embeddings/Quarterly Summary.XLSX')).toBeNull();
    expect(await rendered.file('ppt/charts/_rels/chart1.xml.rels')!.async('text')).not.toMatch(/<Relationship\b/);
    expect(await rendered.file('ppt/charts/chart1.xml')!.async('text')).not.toContain('externalData');
    expect(await rendered.file('[Content_Types].xml')!.async('text')).not.toContain('Quarterly Summary.XLSX');
  });

  test('resolves mixed-case chart part identity and removes workbook and chart metadata', async () => {
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = await JSZip.loadAsync(await fixture());
    zip.file('ppt/charts/Chart1.xml', `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${rel}"><c:externalData r:id="rChart"/></c:chartSpace>`);
    zip.file('ppt/charts/_rels/chart1.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rChart" Type="${rel}/package" Target="../embeddings/Data.xlsx"/></Relationships>`);
    zip.file('ppt/embeddings/Data.xlsx', 'embedded xlsx payload');
    const result = await JSZip.loadAsync(await visibleSlidesCopy(await zip.generateAsync({ type: 'uint8array' })));
    expect(result.file('ppt/embeddings/Data.xlsx')).toBeNull();
    expect(await result.file('ppt/charts/Chart1.xml')!.async('text')).not.toContain('externalData');
    expect(await result.file('ppt/charts/_rels/chart1.xml.rels')!.async('text')).not.toMatch(/<Relationship\b/);
  });

  test('fails closed when a chart relationship has no case-insensitive chart part match', async () => {
    const zip = await JSZip.loadAsync(await fixture());
    zip.file('ppt/charts/_rels/chart1.xml.rels', `<Relationships xmlns="${PKG}"/>`);
    await expect(visibleSlidesCopy(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(/missing chart part/);
  });
});
