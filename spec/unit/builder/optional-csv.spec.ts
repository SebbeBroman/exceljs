import {describe, it, expect} from 'vite-plus/test';
import {workbook, viewWorkbook, readRows} from '../../../excel.js';
import {csv} from '../../../csv.js';

describe('optional CSV entry', () => {
  it('requires the CSV entry for CSV views and row reading', async () => {
    await expect(viewWorkbook('a,b\n1,2')).rejects.toThrow('@sebbebroman/exceljs/csv');
    await expect(readRows('a,b\n1,2')).rejects.toThrow('@sebbebroman/exceljs/csv');
    expect('csv' in workbook()).toBe(false);
  });
  it('selects the first sheet by default and accepts an explicit active-sheet selection', async () => {
    const builder = workbook().sheet('A').row(['first']).sheet('B').row(['second']);
    expect(await csv.stringify(builder)).toBe('first');
    expect(await csv.stringify(builder, {sheetName: 'B'})).toBe('second');
  });
});
