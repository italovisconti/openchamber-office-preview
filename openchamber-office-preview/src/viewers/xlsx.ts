import ExcelJS from 'exceljs';
import * as ssf from 'ssf';
import { throwIfAborted, type ViewerAdapter } from './types';

const PAGE_ROWS = 100;
const PAGE_COLUMNS = 30;

/** Read-only worksheet view. Formula results are never calculated. */
export const mount: ViewerAdapter = async (container, bytes, { signal, onNavigation }) => {
  throwIfAborted(signal);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes.slice().buffer as ArrayBuffer);
  throwIfAborted(signal);
  const worksheets = workbook.worksheets.filter((sheet) => sheet.state === 'visible');
  if (!worksheets.length) throw new Error('Workbook has no visible worksheets');

  const root = document.createElement('section');
  root.className = 'office-xlsx-viewer';
  const nav = document.createElement('nav');
  const select = document.createElement('select');
  select.setAttribute('aria-label', 'Hoja');
  worksheets.forEach((sheet, index) => {
    const option = document.createElement('option'); option.value = String(index); option.textContent = sheet.name; select.append(option);
  });
  nav.append(select);
  const pageInfo = document.createElement('span');
  const prev = document.createElement('button'); prev.type = 'button'; prev.textContent = 'Anterior';
  const next = document.createElement('button'); next.type = 'button'; next.textContent = 'Siguiente';
  nav.append(prev, pageInfo, next);
  const table = document.createElement('table');
  table.setAttribute('aria-label', 'Contenido de la hoja');
  root.append(nav, table);
  container.replaceChildren(root);
  let sheetIndex = 0, page = 0;
  const render = () => {
    const sheet = worksheets[sheetIndex];
    const rows = Math.max(0, sheet.rowCount), cols = Math.max(0, sheet.columnCount);
    const rowPages = Math.max(1, Math.ceil(rows / PAGE_ROWS)), colPages = Math.max(1, Math.ceil(cols / PAGE_COLUMNS));
    page = Math.min(page, rowPages * colPages - 1);
    const rowPage = Math.floor(page / colPages), colPage = page % colPages;
    const startRow = rowPage * PAGE_ROWS + 1, startCol = colPage * PAGE_COLUMNS + 1;
    const body = document.createElement('tbody');
    const endRow = Math.min(startRow + PAGE_ROWS - 1, rows), endCol = Math.min(startCol + PAGE_COLUMNS - 1, cols);
    const merged = new Map<string, { rowspan: number; colspan: number }>();
    const mergedRanges: Array<{ r1: number; r2: number; c1: number; c2: number }> = [];
    for (const range of sheet.model.merges ?? []) {
      const [from, to] = range.split(':');
      const a = decodeAddress(from), b = decodeAddress(to);
      const r1 = Math.max(a.row, startRow), r2 = Math.min(b.row, endRow);
      const c1 = Math.max(a.col, startCol), c2 = Math.min(b.col, endCol);
      if (r1 <= r2 && c1 <= c2) {
        merged.set(`${r1}:${c1}`, { rowspan: r2 - r1 + 1, colspan: c2 - c1 + 1 });
        mergedRanges.push({ r1, r2, c1, c2 });
      }
    }
    for (let r = startRow; r < Math.min(startRow + PAGE_ROWS, rows + 1); r++) {
      const tr = document.createElement('tr');
      for (let c = startCol; c < Math.min(startCol + PAGE_COLUMNS, cols + 1); c++) {
        const info = merged.get(`${r}:${c}`);
        const covered = mergedRanges.some(({ r1, r2, c1, c2 }) => r >= r1 && r <= r2 && c >= c1 && c <= c2 && !(r === r1 && c === c1));
        if (covered && !info) continue;
        const cell = sheet.getCell(r, c), td = document.createElement('td');
        td.textContent = formatCell(cell.value, cell.numFmt, workbook.properties.date1904);
        if (info) { td.rowSpan = info.rowspan; td.colSpan = info.colspan; }
        const fill = cell.fill;
        if (fill?.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor?.argb) {
          const color = safeColor(fill.fgColor.argb); if (color) td.style.backgroundColor = color;
        }
        if (cell.font?.bold) td.style.fontWeight = 'bold';
        if (cell.font?.italic) td.style.fontStyle = 'italic';
        if (cell.font?.color?.argb) {
          const color = safeColor(cell.font.color.argb); if (color) td.style.color = color;
        }
        tr.append(td);
      }
      body.append(tr);
    }
    table.replaceChildren(body);
    pageInfo.textContent = ` ${page + 1}/${rowPages * colPages} `;
    prev.disabled = page === 0; next.disabled = page + 1 >= rowPages * colPages;
    onNavigation?.({ index: sheetIndex, count: worksheets.length, label: sheet.name });
  };
  const onSelect = () => { sheetIndex = Number(select.value); page = 0; render(); };
  const onPrev = () => { if (page) { page--; render(); } };
  const onNext = () => { page++; render(); };
  select.addEventListener('change', onSelect); prev.addEventListener('click', onPrev); next.addEventListener('click', onNext);
  render();
  return { dispose() { select.removeEventListener('change', onSelect); prev.removeEventListener('click', onPrev); next.removeEventListener('click', onNext); root.remove(); } };
};

function formatCell(value: ExcelJS.CellValue, numFmt?: string, date1904 = false): string {
  if (value == null) return '';
  if (typeof value === 'object' && ('formula' in value || 'sharedFormula' in value)) {
    const result = value.result;
    return result == null ? 'Sin resultado guardado' : formatCell(result as ExcelJS.CellValue, numFmt, date1904);
  }
  if (value instanceof Date) {
    try { return new Intl.DateTimeFormat('es', { dateStyle: 'short', timeStyle: 'short' }).format(value); } catch { return value.toISOString(); }
  }
  if (typeof value === 'number' && numFmt) {
    try { return ssf.format(numFmt, value, { date1904 }); } catch { return String(value); }
  }
  if (typeof value === 'object') {
    if ('text' in value) return String(value.text ?? '');
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
    if ('hyperlink' in value) {
      const link = value as { text?: unknown; hyperlink?: unknown };
      return String(link.text ?? link.hyperlink ?? '');
    }
    if ('error' in value) return String(value.error);
  }
  return String(value);
}

function decodeAddress(address: string): { row: number; col: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address);
  if (!match) throw new Error(`Invalid worksheet merge address: ${address}`);
  let col = 0;
  for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]), col };
}

function safeColor(argb: string): string | undefined {
  const hex = argb.replace(/^FF/i, '');
  return /^[0-9a-f]{6}$/i.test(hex) ? `#${hex}` : undefined;
}
