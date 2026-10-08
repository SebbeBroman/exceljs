import {describe, it, expect} from 'vite-plus/test';
import {workbook, load} from '../../../excel.js';

describe('loaded table placement and values', () => {
  it.each([true, false])('rewrites an offset table with headerRow=%s', async headerRow => {
    const rows = [['Ada', 1], ['Bob', null]];
    const bytes = await workbook().sheet('S').table({name: 'People', ref: 'B3', headerRow,
      columns: [{name: 'Name'}, {name: 'Score'}], rows}).writeBuffer();
    const model = await load(bytes);
    expect(model.sheets[0]!.tables![0]).toMatchObject({ref: 'B3', headerRow, rows});
    const target = headerRow ? 'C4' : 'C3';
    const again = await load(await workbook(model).sheet('S').cell(target, 7).writeBuffer());
    expect(again.sheets[0]!.tables![0]!.rows).toEqual([['Ada', 7], ['Bob', null]]);
  });

  it('preserves dates, formulas and cached custom totals', async () => {
    const rows = [[new Date('2020-01-01T00:00:00Z'), {formula: '1+2', result: 3}]];
    const bytes = await workbook().sheet('S').table({name: 'Values', ref: 'D8', totalsRow: true,
      columns: [{name: 'Date', style: {numFmt: 'yyyy-mm-dd'}},
        {name: 'Amount', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(E9)', totalsRowResult: 3}],
      rows}).writeBuffer({useStyles: true});
    const model = await load(bytes);
    const table = model.sheets[0]!.tables![0]!;
    expect(table.rows).toEqual(rows);
    expect(table.columns[1]).toMatchObject({totalsRowFormula: 'SUM(E9)', totalsRowResult: 3});
    const again = await load(await workbook(model).writeBuffer());
    expect(again.sheets[0]!.tables![0]).toEqual(table);
  });
});
