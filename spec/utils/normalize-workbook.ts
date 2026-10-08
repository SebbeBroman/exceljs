import Range from '../../lib/model/range.js';
/** Compare validation coverage across compact-range and older per-cell snapshots. */
export function normalizeWorkbook(value: unknown): string {
  return JSON.stringify(value, (key, item) => {
    if (key === 'worksheet') return undefined;
    // Historical projections omitted table placement/body data. Keep those frozen
    // compatibility hashes while table round-trip tests cover the recovered data.
    if (key === 'tables' && Array.isArray(item)) return item.map(table => {
      const {ref: _ref, rows: _rows, ...rest} = table;
      return {...rest, columns: rest.columns.map((column: Record<string, unknown>) => {
        const {totalsRowFormula: _formula, totalsRowResult: _result, ...original} = column;
        return original;
      })};
    });
    if (key !== 'dataValidations' || !item) return item;
    const cells: Record<string, unknown> = {};
    for (const [address, rule] of Object.entries(item)) {
      const range = new Range(address);
      if (range.count > 100000)
        throw new Error('Compatibility fixtures must not expand large ranges');
      range.forEachAddress(cell => {
        cells[cell] = rule;
      });
    }
    return Object.fromEntries(Object.entries(cells).sort(([a], [b]) => a.localeCompare(b)));
  });
}
