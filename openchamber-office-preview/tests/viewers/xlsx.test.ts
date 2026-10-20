import { describe, expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { mount } from '../../src/viewers/xlsx';

describe('XLSX viewer', () => {
  test('renders formatted values, formula cache status, dates and merged cells in bounded pages', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.properties.date1904 = true;
    const sheet = workbook.addWorksheet('Datos');
    sheet.getCell('A1').value = 0.25; sheet.getCell('A1').numFmt = '0.0%';
    sheet.getCell('A2').value = 2; sheet.getCell('A2').numFmt = '$#,##0.00';
    sheet.getCell('A3').value = { formula: '1+1', result: 2 };
    sheet.getCell('A4').value = { formula: '1+1' };
    sheet.mergeCells('B1:C1'); sheet.getCell('B1').value = 'Merged';
    sheet.getCell('A5').value = new Date('2020-01-01T00:00:00Z');
    for (let i = 6; i <= 155; i++) sheet.getCell(i, 1).value = i;
    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
    const container = document.createElement('div'); document.body.append(container);
    await mount(container as unknown as HTMLElement, bytes, { signal: new AbortController().signal });
    const table = container.querySelector('table')!;
    expect(table.querySelectorAll('tr').length).toBeLessThanOrEqual(100);
    expect(table.textContent).toContain('25.0%');
    expect(table.textContent).toContain('Sin resultado guardado');
    expect(table.querySelector('[colspan="2"]')?.textContent).toBe('Merged');
    container.remove();
  });
});
