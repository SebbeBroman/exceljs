/**
 * Lightweight TypeScript consumer smoke (public API via index.d.ts).
 * Not part of the default Vitest include (see vite.config.js exclude).
 */
import ExcelJS from '../../excel.ts';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2!.getCell('A1').value).to.equal(7);
  });
});
