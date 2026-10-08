import colCache from '../utils/col-cache.js';
import type {CellValue, Style, TableStyleProperties} from './schema.js';

export type TotalsRowFunction =
  | 'none'
  | 'average'
  | 'countNums'
  | 'count'
  | 'max'
  | 'min'
  | 'stdDev'
  | 'var'
  | 'sum'
  | 'custom'
  | string;

export interface TableColumnDef {
  name: string;
  filterButton?: boolean;
  style?: Partial<Style>;
  totalsRowLabel?: string;
  totalsRowFunction?: TotalsRowFunction;
  totalsRowResult?: CellValue;
  totalsRowFormula?: string | null;
}

export interface TableData {
  name: string;
  displayName?: string;
  /** Original typo kept for runtime compatibility with existing data */
  displyName?: string;
  ref: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  style?: TableStyleProperties & {name?: string; theme?: string};
  columns: TableColumnDef[];
  rows: CellValue[][];
  tl?: {row: number; col: number};
  autoFilterRef?: string;
  tableRef?: string;
}

function totalsFormula(table: TableData, column: TableColumnDef): string | null {
  // get the correct formula to apply to the totals row
  switch (column.totalsRowFunction) {
    case 'none':
      return null;
    case 'average':
      return `SUBTOTAL(101,${table.name}[${column.name}])`;
    case 'countNums':
      return `SUBTOTAL(102,${table.name}[${column.name}])`;
    case 'count':
      return `SUBTOTAL(103,${table.name}[${column.name}])`;
    case 'max':
      return `SUBTOTAL(104,${table.name}[${column.name}])`;
    case 'min':
      return `SUBTOTAL(105,${table.name}[${column.name}])`;
    case 'stdDev':
      return `SUBTOTAL(106,${table.name}[${column.name}])`;
    case 'var':
      return `SUBTOTAL(107,${table.name}[${column.name}])`;
    case 'sum':
      return `SUBTOTAL(109,${table.name}[${column.name}])`;
    case 'custom':
      return column.totalsRowFormula ?? null;
    default:
      throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
  }
}

/** Normalize a table and place its headers, values and totals into model cells. */
export function placeTable(
  table: TableData,
  writeCell: (row: number, col: number, value: CellValue, style?: Partial<Style>) => void,
): TableData {
  // set defaults and check is valid
  const assign = <T extends object, K extends keyof T>(o: T, name: K, dflt: T[K]): void => {
    if (o[name] === undefined) {
      o[name] = dflt;
    }
  };
  assign(table, 'headerRow', true);
  assign(table, 'totalsRow', false);

  assign(table, 'style', {});
  assign(table.style!, 'theme', 'TableStyleMedium2');
  assign(table.style!, 'showFirstColumn', false);
  assign(table.style!, 'showLastColumn', false);
  assign(table.style!, 'showRowStripes', false);
  assign(table.style!, 'showColumnStripes', false);

  const assert = (test: unknown, message: string): void => {
    if (!test) {
      throw new Error(message);
    }
  };
  assert(table.ref, 'Table must have ref');
  assert(table.columns, 'Table must have column definitions');
  assert(table.rows, 'Table must have row definitions');

  const decoded = colCache.decodeAddress(table.ref);
  table.tl = {row: decoded.row ?? 0, col: decoded.col ?? 0};
  const {row, col} = table.tl;
  assert(row > 0, 'Table must be on valid row');
  assert(col > 0, 'Table must be on valid col');

  const width = table.columns.length;
  const filterHeight = table.rows.length + (table.headerRow ? 1 : 0);
  const tableHeight = filterHeight + (table.totalsRow ? 1 : 0);

  // autoFilterRef is a range that includes optional headers only
  table.autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);

  // tableRef is a range that includes optional headers and totals
  table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

  table.columns.forEach((column, i) => {
    assert(column.name, `Column ${i} must have a name`);
    // Parsed differential styles carry the format's OOXML model.
    const numFmt = column.style?.numFmt as unknown;
    if (numFmt && typeof numFmt === 'object' && 'formatCode' in numFmt) {
      column.style!.numFmt = String(numFmt.formatCode);
    }
    if (i === 0) {
      assign(column, 'totalsRowLabel', 'Total');
    } else {
      assign(column, 'totalsRowFunction', 'none');
      column.totalsRowFormula = totalsFormula(table, column);
    }
  });

  let count = 0;
  if (table.headerRow) {
    const r = row + count++;
    table.columns.forEach((column, j) => writeCell(r, col + j, column.name, column.style));
  }
  table.rows.forEach(data => {
    const r = row + count++;
    data.forEach((value, j) => writeCell(r, col + j, value, table.columns[j].style));
  });
  if (table.totalsRow) {
    const r = row + count;
    table.columns.forEach((column, j) => {
      const value =
        j === 0
          ? column.totalsRowLabel
          : totalsFormula(table, column)
            ? {
                formula: column.totalsRowFormula!,
                result: column.totalsRowResult as number | string | boolean | Date | undefined,
              }
            : null;
      writeCell(r, col + j, value, column.style);
    });
  }
  return table;
}
