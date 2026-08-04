import colCache from '../utils/col-cache.js';
import type {CellValue, Style, TableStyleProperties} from '../../index.js';

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

export interface TableWorksheet {
  getRow(rowNumber: number): {
    getCell(col: number): {
      value: CellValue;
      style: Partial<Style> & Record<string, unknown>;
    };
  };
}

interface TableCache {
  ref: string;
  width: number;
  tableHeight: number;
}

class Column {
  table: Table;
  column: TableColumnDef;
  index: number;

  constructor(table: Table, column: TableColumnDef, index: number) {
    this.table = table;
    this.column = column;
    this.index = index;
  }

  _set(name: keyof TableColumnDef, value: unknown): void {
    this.table.cacheState();
    (this.column as unknown as Record<string, unknown>)[name] = value;
  }

  get name(): string {
    return this.column.name;
  }
  set name(value: string) {
    this._set('name', value);
  }

  get filterButton(): boolean | undefined {
    return this.column.filterButton;
  }
  set filterButton(value: boolean | undefined) {
    this.column.filterButton = value;
  }

  get style(): Partial<Style> | undefined {
    return this.column.style;
  }
  set style(value: Partial<Style> | undefined) {
    this.column.style = value;
  }

  get totalsRowLabel(): string | undefined {
    return this.column.totalsRowLabel;
  }
  set totalsRowLabel(value: string | undefined) {
    this._set('totalsRowLabel', value);
  }

  get totalsRowFunction(): TotalsRowFunction | undefined {
    return this.column.totalsRowFunction;
  }
  set totalsRowFunction(value: TotalsRowFunction | undefined) {
    this._set('totalsRowFunction', value);
  }

  get totalsRowResult(): CellValue | undefined {
    return this.column.totalsRowResult;
  }
  set totalsRowResult(value: CellValue | undefined) {
    this._set('totalsRowResult', value);
  }

  get totalsRowFormula(): string | null | undefined {
    return this.column.totalsRowFormula;
  }
  set totalsRowFormula(value: string | null | undefined) {
    this._set('totalsRowFormula', value);
  }
}

class Table {
  worksheet!: TableWorksheet;
  table!: TableData;
  private _cache?: TableCache | null;

  constructor(worksheet?: TableWorksheet, table?: TableData) {
    if (worksheet) {
      this.worksheet = worksheet;
    }
    if (table) {
      this.table = table;
      // check things are ok first
      this.validate();

      this.store();
    }
  }

  getFormula(column: TableColumnDef): string | null {
    // get the correct formula to apply to the totals row
    switch (column.totalsRowFunction) {
      case 'none':
        return null;
      case 'average':
        return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
      case 'countNums':
        return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
      case 'count':
        return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
      case 'max':
        return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
      case 'min':
        return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
      case 'stdDev':
        return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
      case 'var':
        return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
      case 'sum':
        return `SUBTOTAL(109,${this.table.name}[${column.name}])`;
      case 'custom':
        return column.totalsRowFormula ?? null;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  get width(): number {
    // width of the table
    return this.table.columns.length;
  }

  get height(): number {
    // height of the table data
    return this.table.rows.length;
  }

  get filterHeight(): number {
    // height of the table data plus optional header row
    return this.height + (this.table.headerRow ? 1 : 0);
  }

  get tableHeight(): number {
    // full height of the table on the sheet
    return this.filterHeight + (this.table.totalsRow ? 1 : 0);
  }

  validate(): void {
    const {table} = this;
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

    const {width, filterHeight, tableHeight} = this;

    // autoFilterRef is a range that includes optional headers only
    table.autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);

    // tableRef is a range that includes optional headers and totals
    table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

    table.columns.forEach((column, i) => {
      assert(column.name, `Column ${i} must have a name`);
      if (i === 0) {
        assign(column, 'totalsRowLabel', 'Total');
      } else {
        assign(column, 'totalsRowFunction', 'none');
        column.totalsRowFormula = this.getFormula(column);
      }
    });
  }

  store(): void {
    // where the table needs to store table data, headers, footers in
    // the sheet...
    const assignStyle = (
      cell: {style: Partial<Style> & Record<string, unknown>},
      style: Partial<Style> | undefined,
    ): void => {
      if (style) {
        Object.keys(style).forEach(key => {
          cell.style[key] = (style as Record<string, unknown>)[key];
        });
      }
    };

    const {worksheet, table} = this;
    const {row, col} = table.tl!;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const {style, name} = column;
        const cell = r.getCell(col + j);
        cell.value = name;
        assignStyle(cell, style);
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;

        assignStyle(cell, table.columns[j].style);
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel as CellValue;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula as string,
              result: column.totalsRowResult as number | string | boolean | Date | undefined,
            };
          } else {
            cell.value = null;
          }
        }

        assignStyle(cell, column.style);
      });
    }
  }

  load(worksheet: TableWorksheet): void {
    // where the table will read necessary features from a loaded sheet
    const {table} = this;
    const {row, col} = table.tl!;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        cell.value = column.name;
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;
      });
    });

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel as CellValue;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula as string,
              result: column.totalsRowResult as number | string | boolean | Date | undefined,
            };
          }
        }
      });
    }
  }

  get model(): TableData {
    return this.table;
  }

  set model(value: TableData) {
    this.table = value;
  }

  // ================================================================
  // TODO: Mutating methods
  cacheState(): void {
    if (!this._cache) {
      this._cache = {
        ref: this.ref,
        width: this.width,
        tableHeight: this.tableHeight,
      };
    }
  }

  commit(): void {
    // changes may have been made that might have on-sheet effects
    if (!this._cache) {
      return;
    }

    // check things are ok first
    this.validate();

    const ref = colCache.decodeAddress(this._cache.ref);
    const refRow = ref.row ?? 0;
    const refCol = ref.col ?? 0;
    if (this.ref !== this._cache.ref) {
      // wipe out whole table footprint at previous location
      for (let i = 0; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(refRow + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(refCol + j);
          cell.value = null;
        }
      }
    } else {
      // clear out below table if it has shrunk
      for (let i = this.tableHeight; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(refRow + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(refCol + j);
          cell.value = null;
        }
      }

      // clear out to right of table if it has lost columns
      for (let i = 0; i < this.tableHeight; i++) {
        const row = this.worksheet.getRow(refRow + i);
        for (let j = this.width; j < this._cache.width; j++) {
          const cell = row.getCell(refCol + j);
          cell.value = null;
        }
      }
    }

    this.store();
  }

  addRow(values: CellValue[], rowNumber?: number): void {
    // Add a row of data, either insert at rowNumber or append
    this.cacheState();

    if (rowNumber === undefined) {
      this.table.rows.push(values);
    } else {
      this.table.rows.splice(rowNumber, 0, values);
    }
  }

  removeRows(rowIndex: number, count = 1): void {
    // Remove a rows of data
    this.cacheState();
    this.table.rows.splice(rowIndex, count);
  }

  getColumn(colIndex: number): Column {
    const column = this.table.columns[colIndex];
    return new Column(this, column, colIndex);
  }

  addColumn(column: TableColumnDef, values: CellValue[], colIndex?: number): void {
    // Add a new column, including column defn and values
    // Inserts at colNumber or adds to the right
    this.cacheState();

    if (colIndex === undefined) {
      this.table.columns.push(column);
      this.table.rows.forEach((row, i) => {
        row.push(values[i]);
      });
    } else {
      this.table.columns.splice(colIndex, 0, column);
      this.table.rows.forEach((row, i) => {
        row.splice(colIndex, 0, values[i]);
      });
    }
  }

  removeColumns(colIndex: number, count = 1): void {
    // Remove a column with data
    this.cacheState();

    this.table.columns.splice(colIndex, count);
    this.table.rows.forEach(row => {
      row.splice(colIndex, count);
    });
  }

  _assign(target: TableData, prop: keyof TableData, value: unknown): void {
    this.cacheState();
    (target as unknown as Record<string, unknown>)[prop] = value;
  }

  get ref(): string {
    return this.table.ref;
  }
  set ref(value: string) {
    this._assign(this.table, 'ref', value);
  }

  get name(): string {
    return this.table.name;
  }
  set name(value: string) {
    this.table.name = value;
  }

  get displayName(): string {
    return this.table.displyName || this.table.name;
  }
  // Original typo `displayNamename` preserved for API compatibility
  set displayNamename(value: string) {
    this.table.displayName = value;
  }

  get headerRow(): boolean | undefined {
    return this.table.headerRow;
  }
  set headerRow(value: boolean) {
    this._assign(this.table, 'headerRow', value);
  }

  get totalsRow(): boolean | undefined {
    return this.table.totalsRow;
  }
  set totalsRow(value: boolean) {
    this._assign(this.table, 'totalsRow', value);
  }

  get theme(): string | undefined {
    return this.table.style?.name;
  }
  set theme(value: string | undefined) {
    if (this.table.style) {
      this.table.style.name = value;
    }
  }

  get showFirstColumn(): boolean | undefined {
    return this.table.style?.showFirstColumn;
  }
  set showFirstColumn(value: boolean | undefined) {
    if (this.table.style) {
      this.table.style.showFirstColumn = value;
    }
  }

  get showLastColumn(): boolean | undefined {
    return this.table.style?.showLastColumn;
  }
  set showLastColumn(value: boolean | undefined) {
    if (this.table.style) {
      this.table.style.showLastColumn = value;
    }
  }

  get showRowStripes(): boolean | undefined {
    return this.table.style?.showRowStripes;
  }
  set showRowStripes(value: boolean | undefined) {
    if (this.table.style) {
      this.table.style.showRowStripes = value;
    }
  }

  get showColumnStripes(): boolean | undefined {
    return this.table.style?.showColumnStripes;
  }
  set showColumnStripes(value: boolean | undefined) {
    if (this.table.style) {
      this.table.style.showColumnStripes = value;
    }
  }
}

export default Table;
export {Table, Column};
