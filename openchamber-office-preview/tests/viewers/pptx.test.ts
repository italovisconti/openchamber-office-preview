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
});
