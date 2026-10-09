import {afterEach, describe, expect, it, vi} from 'vite-plus/test';
import {load, readRows, workbook} from '../../../excel.js';
import CellWriter from '../../../lib/xlsx/xform/sheet/cell-xform.js';
import CellParser from '../../../lib/xlsx/parser/sheet/cell-xform.js';
import StyleParser from '../../../lib/xlsx/parser/style/styles-xform.js';
import StyleWriter from '../../../lib/xlsx/xform/style/styles-xform.js';
import BufferZipWriter from '../../../lib/utils/buffer-zip.js';
import colCache from '../../../lib/utils/col-cache.js';

const work = vi.hoisted(() => ({parsedCells: 0}));
// Wrap the real SAX implementation only in this test module. Callback counts
// distinguish early stopping from parsing everything and slicing the result.
vi.mock('saxen', async importOriginal => {
  const actual = await importOriginal<typeof import('saxen')>();
  return {...actual, Parser: new Proxy(actual.Parser, {
    construct(target, args) {
      const parser = Reflect.construct(target, args) as InstanceType<typeof actual.Parser>;
      type Handler = (...args: unknown[]) => void;
      const on = parser.on.bind(parser) as (event: string, handler: Handler) => void;
      parser.on = ((event: string, handler: Handler) => {
        on(event, event === 'openTag' ? (...args: unknown[]) => {
          if (args[0] === 'c') work.parsedCells++;
          handler(...args);
        } : handler);
      }) as typeof parser.on;
      return parser;
    },
  })};
});

const columns = 8;
const grid = (rows: number) => Array.from({length: rows}, (_, r) =>
  Array.from({length: columns}, (_, c) => r * columns + c + 1),
);

afterEach(() => vi.restoreAllMocks());

// Count expensive boundaries rather than elapsed time. Run small and larger
// fixtures so repeated passes or lost caching cannot hide in a fixed allowance.
describe('deterministic work budgets', () => {
  for (const rows of [32, 128]) {
    it(`writes ${rows} rows with one cell render and one ZIP pass`, async () => {
      const render = vi.spyOn(CellWriter.prototype, 'render');
      const zip = vi.spyOn(BufferZipWriter.prototype, 'toBytes');
      const fullStyles = vi.spyOn(StyleWriter.prototype, 'addStyleModel');
      const bytes = await workbook().sheet('Data').rows(grid(rows)).writeBuffer({useStyles: false});
      expect(render).toHaveBeenCalledTimes(rows * columns);
      expect(zip).toHaveBeenCalledTimes(1);
      expect(fullStyles).not.toHaveBeenCalled();
      // Verify useful output, so bypassing the work cannot satisfy the budget.
      expect(await readRows(bytes, {format: 'xlsx'})).toEqual(grid(rows).map(row => row.map(String)));
    });

    it(`loads ${rows} repeated styled rows without per-cell style resolution`, async () => {
      const bytes = await workbook().sheet('Data').rows(grid(rows))
        .style(`A1:D${rows}`, {numFmt: '0.00'})
        .style(`E1:H${rows}`, {font: {bold: true}}).writeBuffer();
      const styles = vi.spyOn(StyleParser.prototype, 'getStyleModel');
      const classic = vi.spyOn(CellParser.prototype, 'parseOpen');
      const model = await load(bytes);
      expect(model.sheets[0]!.rows).toHaveLength(rows);
      expect(model.sheets[0]!.rows.at(-1)!.cells[8]!.value).toBe(rows * columns);
      expect(styles.mock.calls.length).toBeGreaterThan(0);
      // Two styles, with room for default/row/column lookups, independent of cells.
      expect(styles.mock.calls.length).toBeLessThanOrEqual(4);
      expect(classic).not.toHaveBeenCalled();
    });

    it(`reads a two-row slice of ${rows} rows without visiting later cells`, async () => {
      const bytes = await workbook().sheet('Data').rows(grid(rows)).writeBuffer();
      work.parsedCells = 0;
      const styles = vi.spyOn(StyleParser.prototype, 'getStyleModel');
      const classic = vi.spyOn(CellParser.prototype, 'parseOpen');
      expect(await readRows(bytes, {format: 'xlsx', end: 2})).toEqual(grid(2).map(row => row.map(String)));
      expect(work.parsedCells).toBe(2 * columns);
      expect(styles).not.toHaveBeenCalled();
      expect(classic).not.toHaveBeenCalled();
    });
  }

  it('evicts old address objects while reusing hot addresses', () => {
    const first = colCache.decodeAddress('A1');
    expect(colCache.decodeAddress('A1')).toBe(first);
    for (let row = 1; row <= 100; row++) {
      for (let col = 1; col <= 100; col++) colCache.decodeAddress(`${colCache.n2l(col)}${row}`);
    }
    const latest = colCache.decodeAddress('CV100');
    expect(colCache.decodeAddress('CV100')).toBe(latest);
    expect(colCache.decodeAddress('A1')).not.toBe(first);
    expect(colCache.decodeAddress('A1')).toEqual(first);
  });
});
