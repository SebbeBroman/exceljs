import {describe, it, expect} from 'vite-plus/test';
import {readFile} from 'node:fs/promises';

import {workbook, writeBuffer, load} from '../../../excel.js';
import DataValidationsXform from '../../../lib/xlsx/xform/sheet/data-validations-xform.js';
import Range from '../../../lib/model/range.js';
import {normalizeWorkbook} from '../../utils/normalize-workbook.js';

const rule = {type: 'whole', operator: 'between', formulae: [1, 10]};
function parse(ref: string, xform: DataValidationsXform, formula: string): void {
  xform.parseOpen({name: 'dataValidation', attributes: {type: 'whole', sqref: ref}});
  xform.parseOpen({name: 'formula1', attributes: {}});
  xform.parseText(formula);
  xform.parseClose('formula1');
  xform.parseClose('dataValidation');
}
describe('compact validation ranges', () => {
  it('loads and round-trips whole-grid validation and a defined name without expanding cells', async () => {
    const bytes = await writeBuffer(
      workbook()
        .sheet('S')
        .row([1])
        .dataValidation('A1:XFD1048576', rule as never)
        .definedName('All', 'S!A1:XFD1048576'),
    );
    const first = await load(bytes);
    expect(Object.keys(first.sheets[0].dataValidations!)).toEqual(['A1:XFD1048576']);
    expect(first.definedNames).toEqual([{name: 'All', refersTo: 'S!$A$1:$XFD$1048576'}]);
    const second = await load(await writeBuffer(first));
    expect(second.sheets[0].dataValidations).toEqual(first.sheets[0].dataValidations);
  });
  it('loads the previously stalled issue-1842 fixture', async () => {
    const bytes = await readFile(
      new URL('../../integration/data/test-issue-1842.xlsx', import.meta.url),
    );
    const model = await load(bytes);
    expect(
      model.sheets.some(sheet =>
        Object.keys(sheet.dataValidations ?? {}).includes('A1:XFD1048576'),
      ),
    ).toBe(true);
  });
  it('supports discontiguous references and last-rule-wins overlap through XML re-encoding', async () => {
    const xform = new DataValidationsXform();
    xform.parseOpen({name: 'dataValidations', attributes: {}});
    parse('A1:D4 F1:F3', xform, '1');
    parse('B2:C3', xform, '2');
    const expanded = JSON.parse(normalizeWorkbook({dataValidations: xform.model})).dataValidations;
    expect(Object.keys(expanded)).toHaveLength(19);
    expect(expanded.B2.formulae).toEqual([2]);
    expect(expanded.A1.formulae).toEqual([1]);
    const copy = new DataValidationsXform();
    await copy.parseStream(
      (async function* () {
        yield xform.toXml(xform.model);
      })(),
    );
    expect(normalizeWorkbook({dataValidations: copy.model})).toBe(
      normalizeWorkbook({dataValidations: xform.model}),
    );
  });
  it('preserves coverage and precedence for overlapping random validation rectangles', () => {
    let seed = 29;
    const random = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) % 12) + 1;
    for (let sample = 0; sample < 20; sample++) {
      const xform = new DataValidationsXform();
      xform.parseOpen({name: 'dataValidations', attributes: {}});
      const expected: Record<string, unknown> = {};
      for (let n = 0; n < 15; n++) {
        const range = new Range(random(), random(), random(), random());
        parse(range.shortRange, xform, String(n));
        range.forEachAddress(cell => {
          expected[cell] = {type: 'whole', operator: 'between', formulae: [n]};
        });
      }
      expect(JSON.parse(normalizeWorkbook({dataValidations: xform.model})).dataValidations).toEqual(
        expected,
      );
    }
  });
});
