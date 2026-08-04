import _ from '../utils/under-dash.js';
import Enums from './enums.js';
import colCache from '../utils/col-cache.js';
import type {
  Alignment,
  Borders,
  CellValue,
  Fill,
  Font,
  Protection,
  Style,
} from '../../index.js';
import type Cell from './cell.js';
import type Row from './row.js';

const DEFAULT_COLUMN_WIDTH = 9;

export interface ColumnDefn {
  header?: string | string[];
  key?: string;
  width?: number;
  style?: Partial<Style>;
  hidden?: boolean;
  outlineLevel?: number;
  headers?: string[];
}

export interface ColumnModel {
  min: number;
  max: number;
  width: number;
  style?: Partial<Style>;
  isCustomWidth?: boolean;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

/** Minimal worksheet surface used by Column. */
export interface ColumnWorksheet {
  getCell(row: number, col: number): {value: CellValue};
  getColumnKey(key: string): Column | undefined;
  deleteColumnKey(key: string): void;
  setColumnKey(key: string, column: Column): void;
  properties: {outlineLevelCol?: number};
  eachRow(
    options: {includeEmpty?: boolean} | ((row: Row, rowNumber: number) => void) | null,
    iteratee?: (row: Row, rowNumber: number) => void,
  ): void;
}

// Column defines the column properties for 1 column.
// This includes header rows, widths, key, (style), etc.
// Worksheet will condense the columns as appropriate during serialization
class Column {
  _worksheet: ColumnWorksheet;
  _number: number;
  _header?: string | string[];
  _key?: string;
  width?: number;
  style!: Partial<Style> & Record<string, unknown>;
  _hidden?: boolean;
  _outlineLevel?: number;

  constructor(worksheet: ColumnWorksheet, number: number, defn?: ColumnDefn | false | null) {
    this._worksheet = worksheet;
    this._number = number;
    if (defn !== false) {
      // sometimes defn will follow
      this.defn = defn as ColumnDefn | null | undefined;
    }
  }

  get number(): number {
    return this._number;
  }

  get worksheet(): ColumnWorksheet {
    return this._worksheet;
  }

  get letter(): string {
    return colCache.n2l(this._number);
  }

  get isCustomWidth(): boolean {
    return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
  }

  get defn(): ColumnDefn {
    return {
      header: this._header,
      key: this.key,
      width: this.width,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
    };
  }

  set defn(value: ColumnDefn | null | undefined) {
    if (value) {
      this.key = value.key;
      this.width = value.width !== undefined ? value.width : DEFAULT_COLUMN_WIDTH;
      this.outlineLevel = value.outlineLevel ?? 0;
      if (value.style) {
        this.style = value.style as Partial<Style> & Record<string, unknown>;
      } else {
        this.style = {};
      }

      // headers must be set after style
      this.header = value.header;
      this._hidden = !!value.hidden;
    } else {
      delete this._header;
      delete this._key;
      delete this.width;
      this.style = {};
      this.outlineLevel = 0;
    }
  }

  get headers(): Array<string | undefined> {
    return this._header && this._header instanceof Array ? this._header : [this._header];
  }

  get header(): string | string[] | undefined {
    return this._header;
  }

  set header(value: string | string[] | undefined) {
    if (value !== undefined) {
      this._header = value;
      this.headers.forEach((text, index) => {
        this._worksheet.getCell(index + 1, this.number).value = text as CellValue;
      });
    } else {
      this._header = undefined;
    }
  }

  get key(): string | undefined {
    return this._key;
  }

  set key(value: string | undefined) {
    const column = this._key && this._worksheet.getColumnKey(this._key);
    if (column === this) {
      this._worksheet.deleteColumnKey(this._key!);
    }

    this._key = value;
    if (value) {
      this._worksheet.setColumnKey(this._key!, this);
    }
  }

  get hidden(): boolean {
    return !!this._hidden;
  }

  set hidden(value: boolean) {
    this._hidden = value;
  }

  get outlineLevel(): number {
    return this._outlineLevel || 0;
  }

  set outlineLevel(value: number) {
    this._outlineLevel = value;
  }

  get collapsed(): boolean {
    return !!(
      this._outlineLevel && this._outlineLevel >= (this._worksheet.properties.outlineLevelCol ?? 0)
    );
  }

  toString(): string {
    return JSON.stringify({
      key: this.key,
      width: this.width,
      headers: this.headers.length ? this.headers : undefined,
    });
  }

  equivalentTo(other: Column): boolean {
    return (
      this.width === other.width &&
      this.hidden === other.hidden &&
      this.outlineLevel === other.outlineLevel &&
      _.isEqual(this.style, other.style)
    );
  }

  get isDefault(): boolean {
    if (this.isCustomWidth) {
      return false;
    }
    if (this.hidden) {
      return false;
    }
    if (this.outlineLevel) {
      return false;
    }
    const s = this.style;
    if (s && (s.font || s.numFmt || s.alignment || s.border || s.fill || s.protection)) {
      return false;
    }
    return true;
  }

  get headerCount(): number {
    return this.headers.length;
  }

  eachCell(
    options: {includeEmpty?: boolean} | ((cell: Cell, rowNumber: number) => void),
    iteratee?: (cell: Cell, rowNumber: number) => void,
  ): void {
    const colNumber = this.number;
    let opts: {includeEmpty?: boolean} | null = options as {includeEmpty?: boolean};
    let fn = iteratee;
    if (!fn) {
      fn = options as (cell: Cell, rowNumber: number) => void;
      opts = null;
    }
    this._worksheet.eachRow(opts, (row, rowNumber) => {
      fn!(row.getCell(colNumber), rowNumber);
    });
  }

  get values(): CellValue[] {
    const v: CellValue[] = [];
    this.eachCell((cell, rowNumber) => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        v[rowNumber] = cell.value;
      }
    });
    return v;
  }

  set values(v: CellValue[] | null | undefined) {
    if (!v) {
      return;
    }
    const colNumber = this.number;
    let offset = 0;
    if (Object.hasOwn(v, '0')) {
      // assume contiguous array, start at row 1
      offset = 1;
    }
    v.forEach((value, index) => {
      this._worksheet.getCell(index + offset, colNumber).value = value;
    });
  }

  // =========================================================================
  // styles
  _applyStyle(name: string, value: unknown): unknown {
    this.style[name] = value;
    this.eachCell(cell => {
      (cell as unknown as Record<string, unknown>)[name] = value;
    });
    return value;
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }

  set numFmt(value: string | undefined) {
    this._applyStyle('numFmt', value);
  }

  get font(): Font | undefined {
    return this.style.font as Font | undefined;
  }

  set font(value: Font | undefined) {
    this._applyStyle('font', value);
  }

  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    this._applyStyle('alignment', value);
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection;
  }

  set protection(value: Partial<Protection> | undefined) {
    this._applyStyle('protection', value);
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }

  set border(value: Partial<Borders> | undefined) {
    this._applyStyle('border', value);
  }

  get fill(): Fill | undefined {
    return this.style.fill;
  }

  set fill(value: Fill | undefined) {
    this._applyStyle('fill', value);
  }

  // =============================================================================
  // static functions

  static toModel(columns: Column[] | null | undefined): ColumnModel[] | undefined {
    // Convert array of Column into compressed list cols
    const cols: ColumnModel[] = [];
    let col: ColumnModel | null = null;
    if (columns) {
      columns.forEach((column, index) => {
        if (column.isDefault) {
          if (col) {
            col = null;
          }
        } else if (!col || !column.equivalentTo(col as unknown as Column)) {
          // original compares to previous Column instance stored in `col` which is a model after first push
          // Runtime: after first assignment, col is a model object, so equivalentTo is called with model as "other"
          // Keep same structure:
          col = {
            min: index + 1,
            max: index + 1,
            width: column.width !== undefined ? column.width : DEFAULT_COLUMN_WIDTH,
            style: column.style,
            isCustomWidth: column.isCustomWidth,
            hidden: column.hidden,
            outlineLevel: column.outlineLevel,
            collapsed: column.collapsed,
          };
          cols.push(col);
        } else {
          col.max = index + 1;
        }
      });
    }
    return cols.length ? cols : undefined;
  }

  static fromModel(worksheet: ColumnWorksheet, cols?: ColumnModel[] | null): Column[] | null {
    cols = cols || [];
    const columns: Column[] = [];
    let count = 1;
    let index = 0;
    /**
     * sort cols by min
     * If it is not sorted, the subsequent column configuration will be overwritten
     * */
    cols = cols.slice().sort(function (pre, next) {
      return pre.min - next.min;
    });
    while (index < cols.length) {
      const col = cols[index++];
      while (count < col.min) {
        columns.push(new Column(worksheet, count++));
      }
      while (count <= col.max) {
        columns.push(new Column(worksheet, count++, col));
      }
    }
    return columns.length ? columns : null;
  }
}

export default Column;
export {Column};
