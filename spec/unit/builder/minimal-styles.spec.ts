import {describe, it, expect} from 'vite-plus/test';
import {workbook, load} from '../../../excel.js';
import {unzipSync, strFromU8} from 'fflate';

describe('minimal styles exports', () => {
  it('keeps dates typed without enabling cell styling', async () => {
    const date = new Date('2021-02-03T00:00:00Z');
    const bytes = await workbook().sheet('S').rows([[1, date]]).writeBuffer();
    const again = await load(bytes);
    expect(again.sheets[0]!.rows[0]!.cells[2]!.value).toEqual(date);
  });

  it('keeps differential formats with cell styling disabled', async () => {
    const bytes = await workbook().sheet('S')
      .table({name: 'Values', ref: 'A1', columns: [{name: 'Value', style: {numFmt: '0.0000'}}], rows: [[3]]})
      .conditionalFormatting({ref: 'A2', rules: [{type: 'cellIs', priority: 1, operator: 'greaterThan', formulae: [1], style: {font: {bold: true}, numFmt: '0.0000'}}]})
      .writeBuffer({useStyles: false});
    const files = unzipSync(bytes);
    const styles = strFromU8(files['xl/styles.xml']!);
    expect(styles).toContain('formatCode="0.0000"');
    expect(styles).toContain('<dxfs count="2">');
    const model = await load(bytes);
    expect(model.sheets[0]!.conditionalFormattings![0]!.rules[0]!.style!.font!.bold).toBe(true);
    expect(model.sheets[0]!.tables![0]!.columns[0]!.style!.numFmt).toMatchObject({formatCode: '0.0000'});
  });
});
