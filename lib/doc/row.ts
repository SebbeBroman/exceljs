import _ from '../utils/under-dash.js';
import Enums from './enums.js';
import colCache from '../utils/col-cache.js';
import Cell from './cell.js';
import type {
  Alignment,
  Borders,
  CellValue,
  Fill,
  Font,
  Protection,
  RowBreak,
  Style,
} from '../../index.js';
import type Column from './column.js';
import type {ValueTypeCode} from './enums.js';

export interface RowDimensions {
  min: number;
  max: number;
}

export interface RowCellAddress {
  address: string;
  row: number;
  col: number;
  $col$row?: string;
}

export interface RowModelCell {
  type: ValueTypeCode | number;
  address?: string;
  [key: string]: unknown;
}

export interface RowModelData {
  cells: RowModelCell[];
  number: number;
  min: number;
  max: number;
  height?: number;
  style?: Partial<Style>;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

/** Minimal worksheet surface used by Row. */
export interface RowWorksheet {
  getColumn(col: number): Column;
  getColumnKey(key: string): Column | undefined;
  eachColumnKey(callback: (column: Column, key: string) => void): void;
  _commitRow(row: Row): void;
  rowBreaks: RowBreak[];
  properties: {outlineLevelRow?: number};
}

class Row {
  _worksheet!: RowWorksheet;
  _number: number;
  _cells: Array<Cell | undefined>;
  style: Partial<Style> & Record<string, unknown>;
  _outlineLevel?: number;
  _hidden?: boolean;
  height?: number;

  constructor(worksheet: RowWorksheet, number: number) {
    this._worksheet = worksheet;
    this._number = number;
    this._cells = [];
    this.style = {};
    this.outlineLevel = 0;
  }

  // return the row number
  get number(): number {
    return this._number;
  }

  get worksheet(): RowWorksheet {
    return this._worksheet;
  }

  // Inform Streaming Writer that this row (and all rows before it) are complete
  // and ready to write. Has no effect on Worksheet document
  commit(): void {
    this._worksheet._commitRow(this);
  }

  // helps GC by breaking cyclic references
  destroy(): void {
    delete (this as Partial<Row>)._worksheet;
    delete (this as Partial<Row>)._cells;
    delete (this as Partial<Row>).style;
  }

  findCell(colNumber: number): Cell | undefined {
    return this._cells[colNumber - 1];
  }

  // given {address, row, col}, find or create new cell
  getCellEx(address: RowCellAddress): Cell {
    let cell = this._cells[address.col - 1];
    if (!cell) {
      const column = this._worksheet.getColumn(address.col);
      cell = new Cell(this, column, address.address);
      this._cells[address.col - 1] = cell;
    }
    return cell;
  }

  // get cell by key, letter or column number
  getCell(col: number | string): Cell {
    if (typeof col === 'string') {
      // is it a key?
      const column = this._worksheet.getColumnKey(col);
      if (column) {
        col = column.number;
      } else {
        col = colCache.l2n(col);
      }
    }
    return (
      this._cells[col - 1] ||
      this.getCellEx({
        address: colCache.encodeAddress(this._number, col),
        row: this._number,
        col,
      })
    );
  }

  // remove cell(s) and shift all higher cells down by count
  splice(start: number, count: number, ...inserts: CellValue[]): void {
    const nKeep = start + count;
    const nExpand = inserts.length - count;
    const nEnd = this._cells.length;
    let i: number;
    let cSrc: Cell | undefined;
    let cDst: Cell | undefined;

    if (nExpand < 0) {
      // remove cells
      for (i = start + inserts.length; i <= nEnd; i++) {
        cDst = this._cells[i - 1];
        cSrc = this._cells[i - nExpand - 1];
        if (cSrc) {
          cDst = this.getCell(i);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;

          cDst._comment = cSrc._comment;
        } else if (cDst) {
          cDst.value = null;
          cDst.style = {};

          cDst._comment = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        cSrc = this._cells[i - 1];
        if (cSrc) {
          cDst = this.getCell(i + nExpand);
          cDst.value = cSrc.value;
          cDst.style = cSrc.style;

          cDst._comment = cSrc._comment;
        } else {
          this._cells[i + nExpand - 1] = undefined;
        }
      }
    }

    // now add the new values
    for (i = 0; i < inserts.length; i++) {
      cDst = this.getCell(start + i);
      cDst.value = inserts[i];
      cDst.style = {};

      cDst._comment = undefined;
    }
  }

  // Iterate over all non-null cells in this row
  eachCell(
    options: {includeEmpty?: boolean} | ((cell: Cell, colNumber: number) => void),
    iteratee?: (cell: Cell, colNumber: number) => void,
  ): void {
    let opts: {includeEmpty?: boolean} | null = options as {includeEmpty?: boolean};
    let fn = iteratee;
    if (!fn) {
      fn = options as (cell: Cell, colNumber: number) => void;
      opts = null;
    }
    if (opts && opts.includeEmpty) {
      const n = this._cells.length;
      for (let i = 1; i <= n; i++) {
        fn(this.getCell(i), i);
      }
    } else {
      this._cells.forEach((cell, index) => {
        if (cell && cell.type !== Enums.ValueType.Null) {
          fn!(cell, index + 1);
        }
      });
    }
  }

  // ===========================================================================
  // Page Breaks
  addPageBreak(lft?: number, rght?: number): void {
    const ws = this._worksheet;
    const left = Math.max(0, (lft ?? 0) - 1) || 0;
    const right = Math.max(0, (rght ?? 0) - 1) || 16838;
    const pb: RowBreak = {
      id: this._number,
      max: right,
      min: left || 0,
      man: 1,
    };
    if (left) pb.min = left;

    ws.rowBreaks.push(pb);
  }

  // return a sparse array of cell values
  get values(): CellValue[] {
    const values: CellValue[] = [];
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        values[cell.col] = cell.value;
      }
    });
    return values;
  }

  // set the values by contiguous or sparse array, or by key'd object literal
  set values(value: CellValue[] | Record<string, CellValue> | null | undefined) {
    // this operation is not additive - any prior cells are removed
    this._cells = [];
    if (!value) {
      // empty row
    } else if (value instanceof Array) {
      let offset = 0;
      if (Object.hasOwn(value, '0')) {
        // contiguous array - start at column 1
        offset = 1;
      }
      value.forEach((item, index) => {
        if (item !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, index + offset),
            row: this._number,
            col: index + offset,
          }).value = item;
        }
      });
    } else {
      // assume object with column keys
      const obj = value as Record<string, CellValue>;
      this._worksheet.eachColumnKey((column, key) => {
        if (obj[key] !== undefined) {
          this.getCellEx({
            address: colCache.encodeAddress(this._number, column.number),
            row: this._number,
            col: column.number,
          }).value = obj[key];
        }
      });
    }
  }

  // returns true if the row includes at least one cell with a value
  get hasValues(): boolean {
    return _.some(this._cells, (cell: Cell | undefined) => cell && cell.type !== Enums.ValueType.Null);
  }

  get cellCount(): number {
    return this._cells.length;
  }

  get actualCellCount(): number {
    let count = 0;
    this.eachCell(() => {
      count++;
    });
    return count;
  }

  // get the min and max column number for the non-null cells in this row or null
  get dimensions(): RowDimensions | null {
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell && cell.type !== Enums.ValueType.Null) {
        if (!min || min > cell.col) {
          min = cell.col;
        }
        if (max < cell.col) {
          max = cell.col;
        }
      }
    });
    return min > 0
      ? {
          min,
          max,
        }
      : null;
  }

  // =========================================================================
  // styles
  _applyStyle(name: string, value: unknown): unknown {
    this.style[name] = value;
    this._cells.forEach(cell => {
      if (cell) {
        (cell as unknown as Record<string, unknown>)[name] = value;
      }
    });
    return value;
  }

  get numFmt(): string | undefined {
    return this.style.numFmt as string | undefined;
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
    return this.style.alignment as Partial<Alignment> | undefined;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    this._applyStyle('alignment', value);
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection as Partial<Protection> | undefined;
  }

  set protection(value: Partial<Protection> | undefined) {
    this._applyStyle('protection', value);
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border as Partial<Borders> | undefined;
  }

  set border(value: Partial<Borders> | undefined) {
    this._applyStyle('border', value);
  }

  get fill(): Fill | undefined {
    return this.style.fill as Fill | undefined;
  }

  set fill(value: Fill | undefined) {
    this._applyStyle('fill', value);
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
      this._outlineLevel && this._outlineLevel >= (this._worksheet.properties.outlineLevelRow ?? 0)
    );
  }

  // =========================================================================
  get model(): RowModelData | null {
    const cells: unknown[] = [];
    let min = 0;
    let max = 0;
    this._cells.forEach(cell => {
      if (cell) {
        const cellModel = cell.model;
        if (cellModel) {
          if (!min || min > cell.col) {
            min = cell.col;
          }
          if (max < cell.col) {
            max = cell.col;
          }
          cells.push(cellModel);
        }
      }
    });

    return this.height || cells.length
      ? {
          cells: cells as RowModelCell[],
          number: this.number,
          min,
          max,
          height: this.height,
          style: this.style,
          hidden: this.hidden,
          outlineLevel: this.outlineLevel,
          collapsed: this.collapsed,
        }
      : null;
  }

  set model(value: RowModelData) {
    if (value.number !== this._number) {
      throw new Error('Invalid row number in model');
    }
    this._cells = [];
    let previousAddress: RowCellAddress | undefined;
    value.cells.forEach(cellModel => {
      switch (cellModel.type) {
        case Cell.Types.Merge:
          // special case - don't add this types
          break;
        default: {
          let address: RowCellAddress | undefined;
          if (cellModel.address) {
            const decoded = colCache.decodeAddress(cellModel.address as string);
            address = {
              address: decoded.address,
              row: decoded.row ?? 0,
              col: decoded.col ?? 0,
              $col$row: decoded.$col$row,
            };
          } else if (previousAddress) {
            // This is a <c> element without an r attribute
            // Assume that it's the cell for the next column
            const {row} = previousAddress;
            const col = previousAddress.col + 1;
            address = {
              row,
              col,
              address: colCache.encodeAddress(row, col),
              $col$row: `$${colCache.n2l(col)}$${row}`,
            };
          }
          previousAddress = address;
          const cell = this.getCellEx(address!);
          cell.model = cellModel as never;
          break;
        }
      }
    });

    if (value.height) {
      this.height = value.height;
    } else {
      delete this.height;
    }

    this.hidden = !!value.hidden;
    this.outlineLevel = value.outlineLevel || 0;

    this.style = (value.style && JSON.parse(JSON.stringify(value.style))) || {};
  }
}

export default Row;
export {Row};
