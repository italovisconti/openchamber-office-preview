import JSZip from 'jszip';

const MAX_SLIDES = 500;
const NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

async function requiredXml(zip: JSZip, path: string): Promise<Document> {
  const file = zip.file(path);
  if (!file) throw new Error(`PPTX missing required part: ${path}`);
  const doc = new DOMParser().parseFromString(await file.async('text'), 'application/xml');
  if (doc.querySelector('parsererror')) throw new Error(`Invalid XML part: ${path}`);
  return doc;
}
function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/') || /^[a-z]+:/i.test(target)) throw new Error('Unsupported external/absolute PPTX relationship');
  const parts = source.split('/'); parts.pop();
  for (const part of target.split('/')) { if (!part || part === '.') continue; if (part === '..') parts.pop(); else parts.push(part); }
  return parts.join('/');
}
function relPath(part: string): string { const i = part.lastIndexOf('/'); return `${i < 0 ? '' : part.slice(0, i + 1)}_rels/${i < 0 ? part : part.slice(i + 1)}.rels`; }

/** Produce a disposable rendering copy; never mutates the caller's bytes. */
export async function visibleSlidesCopy(bytes: Uint8Array): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(bytes.slice(), { checkCRC32: true });
  const presentation = await requiredXml(zip, 'ppt/presentation.xml');
  const relsDoc = await requiredXml(zip, 'ppt/_rels/presentation.xml.rels');
  const rels = new Map([...relsDoc.getElementsByTagName('*')].filter((r) => r.localName === 'Relationship').map((r) => [r.getAttribute('Id')!, r]));
  const list = [...presentation.getElementsByTagName('*')].find((el) => el.localName === 'sldIdLst');
  if (!list) throw new Error('PPTX missing p:sldIdLst');
  const slideIds = [...list.getElementsByTagName('*')].filter((el) => el.localName === 'sldId');
  if (!slideIds.length || slideIds.length > MAX_SLIDES) throw new Error(`PPTX slide count must be between 1 and ${MAX_SLIDES}`);
  const kept = new Set<string>();
  for (const id of slideIds) {
    const relationshipId = id.getAttributeNS(REL, 'id') ?? id.getAttribute('r:id');
    const rel = relationshipId ? rels.get(relationshipId) : undefined;
    if (!rel) throw new Error('PPTX slide relationship is missing');
    if (rel.getAttribute('TargetMode') === 'External') throw new Error('External PPTX relationships are not supported');
    const part = resolveTarget('ppt/presentation.xml', rel.getAttribute('Target') ?? '');
    const slide = await requiredXml(zip, part);
    const show = slide.documentElement.getAttribute('show');
    if (show === '0' || show === 'false') { zip.remove(part); zip.remove(relPath(part)); id.remove(); rel.remove(); }
    else kept.add(part);
  }
  if (!kept.size) throw new Error('PPTX has no visible slides');
  // Remove orphan slide parts and relationships so the renderer cannot discover hidden content by package scan.
  for (const name of Object.keys(zip.files)) if (/^ppt\/slides\/slide\d+\.xml$/.test(name) && !kept.has(name)) { zip.remove(name); zip.remove(relPath(name)); }
  // Cached chart XML is sufficient; remove the linked XLSX before invoking the renderer.
  for (const name of Object.keys(zip.files)) {
    const match = /^(ppt\/charts)\/_rels\/(chart\d+\.xml)\.rels$/i.exec(name);
    if (!match) continue;
    const chartIdentity = `${match[1]}/${match[2]}`.toLowerCase();
    const chartPath = Object.keys(zip.files).find((path) => !zip.files[path].dir && path.toLowerCase() === chartIdentity);
    if (!chartPath) throw new Error('PPTX chart relationship references a missing chart part');
    const chartFile = zip.file(chartPath), relsFile = zip.file(name);
    if (!chartFile || !relsFile) throw new Error('PPTX chart relationship part is missing');
    const chart = new DOMParser().parseFromString(await chartFile.async('text'), 'application/xml');
    const rels = new DOMParser().parseFromString(await relsFile.async('text'), 'application/xml');
    if (chart.querySelector('parsererror') || rels.querySelector('parsererror')) throw new Error('Invalid chart XML');
    let changed = false;
    for (const rel of [...rels.getElementsByTagNameNS(PKG_REL, 'Relationship')]) {
      const mode = rel.getAttribute('TargetMode');
      if (rel.getAttribute('Type') !== `${REL}/package`) continue;
      if (mode !== null && mode !== 'Internal') throw new Error('External PPTX chart workbook relationships are not supported');
      const target = resolveChartWorkbookTarget(name, rel.getAttribute('Target') ?? '');
      if (!target) throw new Error('Invalid PPTX chart workbook relationship target');
      const id = rel.getAttribute('Id');
      if (!id) throw new Error('PPTX chart workbook relationship is missing its identifier');
      const externalData = [...chart.getElementsByTagNameNS(CHART_NS, 'externalData')].find((el) => el.getAttributeNS(REL, 'id') === id);
      if (!externalData) throw new Error('PPTX chart workbook relationship has no matching externalData reference');
      const actualTarget = Object.keys(zip.files).find((path) => !zip.files[path].dir && path.toLowerCase() === target.toLowerCase());
      if (!actualTarget) throw new Error('PPTX chart workbook part is missing from the rendering copy');
      externalData.parentNode?.removeChild(externalData);
      rel.parentNode?.removeChild(rel);
      zip.remove(actualTarget);
      changed = true;
    }
    if (changed) {
      zip.file(chartPath, new XMLSerializer().serializeToString(chart));
      zip.file(name, new XMLSerializer().serializeToString(rels));
    }
  }
  zip.file('ppt/presentation.xml', new XMLSerializer().serializeToString(presentation));
  zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
  // The engine also discovers slides through content types. Removed parts must
  // not remain advertised there, including hidden slides and chart workbooks.
  const contentTypes = await requiredXml(zip, '[Content_Types].xml');
  for (const entry of [...contentTypes.getElementsByTagNameNS('*', 'Override')]) {
    const part = entry.getAttribute('PartName')?.replace(/^\//, '');
    if (part && !zip.file(part)) entry.remove();
  }
  zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(contentTypes));
  const out = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  return out;
}
function resolveChartWorkbookTarget(sourceRels: string, target: string): string | null {
  if (!/^ppt\/charts\/_rels\/chart\d+\.xml\.rels$/i.test(sourceRels)) return null;
  const match = /^\.\.\/embeddings\/([^/\\?#%]+\.xlsx)$/i.exec(target);
  if (!match || /[\u0000-\u001f:\\/?#%]/.test(match[1]) || match[1] === '..' || match[1] === '.') return null;
  return `ppt/embeddings/${match[1]}`;
}
