import Enums from './enums.js';
import colCache from '../utils/col-cache.js';
import Cell from './cell.js';
import Note from './note.js';
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

/**
 * Compact (flat) cell storage: no Cell / Value strategy objects.
 * Materialized to Cell on getCell / findCell / eachCell.
 * Write path (row.model) reads compact slots without hydrating.
 */
export interface CompactCell {
  /** Brand: distinguishes from Cell without instanceof on hot paths. */
  readonly _c: 1;
  col: number;
  address: string;
  /** ValueType code; -1 for JSON-stringified plain objects. */
  type: number;
  value: unknown;
  /** Prebuilt xform model (without style/comment); avoids rebuild on row.model. */
  wm: RowModelCell;
  style?: Partial<Style> & Record<string, unknown>;
  comment?: Note;
}

/** Non-materializing cell slot view for plain projection / notes. */
export interface CellValueInfo {
  col: number;
  address: string;
  type: number;
  value: unknown;
  style?: Partial<Style> & Record<string, unknown>;
  /** Public note payload (string or Comment), if any. */
  note?: unknown;
}

export type CellSlot = Cell | CompactCell;

function isCompact(entry: CellSlot | undefined | null): entry is CompactCell {
  return entry != null && (entry as CompactCell)._c === 1;
}

function styleHasKeys(style: (Partial<Style> & Record<string, unknown>) | undefined): boolean {
  if (!style) return false;
  // faster than Object.keys alloc for the common empty-style case
  for (const _k in style) return true;
  return false;
}

/** Attach style/comment onto the prebuilt write model for xform. */
function compactToModel(entry: CompactCell): RowModelCell {
  const model = entry.wm;
  if (entry.style) model.style = entry.style;
  else delete model.style;
  if (entry.comment) {
    model.comment = (entry.comment as {model?: unknown}).model ?? entry.comment;
  }
  return model;
}

function buildWriteModel(address: string, type: number, value: unknown): RowModelCell {
  switch (type) {
    case Enums.ValueType.Null:
      return {address, type};
    case Enums.ValueType.Number:
    case Enums.ValueType.String:
    case Enums.ValueType.Date:
    case Enums.ValueType.Boolean:
      return {address, type, value};
    default:
      return Cell.valueToModel(address, value, type) as RowModelCell;
  }
}

/**
 * Reconstruct the public CellValue from a post-reconcile xform cell model.
 * Used for CompactCell.value so materialize / eachValue work without a Cell.
 */
function valueFromCellModel(cellModel: RowModelCell): unknown {
  const type = cellModel.type as number;
  switch (type) {
    case Enums.ValueType.Null:
      return null;
    case Enums.ValueType.Number:
    case Enums.ValueType.String:
    case Enums.ValueType.Date:
    case Enums.ValueType.Boolean:
    case Enums.ValueType.Error:
    case Enums.ValueType.SharedString:
    case Enums.ValueType.RichText:
      // JSONValue models carry rawValue (rare on load path)
      if (cellModel.rawValue !== undefined) return cellModel.rawValue;
      return cellModel.value;
    case Enums.ValueType.Hyperlink: {
      const v: {
        text?: unknown;
        hyperlink?: unknown;
        tooltip?: unknown;
        formula?: unknown;
        sharedFormula?: unknown;
        shareType?: unknown;
        ref?: unknown;
      } = {
        text: cellModel.text,
        hyperlink: cellModel.hyperlink,
      };
      if (cellModel.tooltip != null) v.tooltip = cellModel.tooltip;
      // A formula cell reconciled to a hyperlink keeps its `formula` (the
      // reconcile moves result -> text). Carry the formula fields through so
      // `cell.model.formula` survives, matching the direct-model load path.
      if (cellModel.formula != null) v.formula = cellModel.formula;
      if (cellModel.sharedFormula != null) v.sharedFormula = cellModel.sharedFormula;
      if (cellModel.shareType != null) v.shareType = cellModel.shareType;
      if (cellModel.ref != null) v.ref = cellModel.ref;
      return v;
    }
    case Enums.ValueType.Formula: {
      const v: {
        formula?: unknown;
        sharedFormula?: unknown;
        shareType?: unknown;
        ref?: unknown;
        result?: unknown;
      } = {};
      if (cellModel.formula != null) v.formula = cellModel.formula;
      if (cellModel.sharedFormula != null) v.sharedFormula = cellModel.sharedFormula;
      if (cellModel.shareType != null) v.shareType = cellModel.shareType;
      if (cellModel.ref != null) v.ref = cellModel.ref;
      if (cellModel.result !== undefined) v.result = cellModel.result;
      return v;
    }
    default:
      return cellModel.value;
  }
}

/** Types safe to keep as CompactCell on load (no Cell / Value strategy). */
function isCompactLoadType(type: number): boolean {
  switch (type) {
    case Enums.ValueType.Null:
    case Enums.ValueType.Number:
    case Enums.ValueType.String:
    case Enums.ValueType.Date:
    case Enums.ValueType.Boolean:
    case Enums.ValueType.Hyperlink:
    case Enums.ValueType.Formula:
    case Enums.ValueType.SharedString:
    case Enums.ValueType.RichText:
    case Enums.ValueType.Error:
      return true;
    default:
      return false;
  }
}

class Row {
  _worksheet!: RowWorksheet;
  _number: number;
  /** Sparse slots: CompactCell until API needs a full Cell. */
  _cells: Array<CellSlot | undefined>;
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

  /** Materialize a compact slot into a full Cell (in place). */
  _materialize(col: number, compact: CompactCell): Cell {
    const column = this._worksheet.getColumn(col);
    const cell = new Cell(this, column, compact.address, {
      validateAddress: false,
      value: compact.value,
      style: compact.style || Cell.mergeStyles(this.style, column.style, {}),
    });
    if (compact.comment) {
      cell._comment = compact.comment as never;
    }
    this._cells[col - 1] = cell;
    return cell;
  }

  /**
   * Store a value as a compact slot (no Cell allocation).
   * @param column optional pre-resolved column (avoids getColumn on key'd addRow)
   */
  _setCompact(col: number, value: unknown, column?: Column): void {
    if (value === undefined) {
      this._cells[col - 1] = undefined;
      return;
    }

    const rawType = Cell.getValueType(value);
    const type = rawType === undefined ? -1 : rawType;
    const address = colCache.encodeAddress(this._number, col);
    const colObj = column || this._worksheet.getColumn(col);
    const compact: CompactCell = {
      _c: 1,
      col,
      address,
      type,
      value: value as CellValue,
      wm: buildWriteModel(address, type, value),
    };
    // Only allocate a style object when row or column actually carries styles
    if (styleHasKeys(this.style) || styleHasKeys(colObj.style)) {
      const style = Cell.mergeStyles(this.style, colObj.style, {});
      if (styleHasKeys(style)) {
        compact.style = style;
      }
    }
    this._cells[col - 1] = compact;
  }

  /**
   * Store a post-reconcile xform cell model as a CompactCell (load path).
   * Falls back to full Cell for types that cannot be represented compactly.
   *
   * The reconciled model is reused in place as the write model (`wm`): it is
   * garbage after hydration, so instead of copying every key into a fresh
   * object (one alloc + N copies per cell), style/comment are extracted and
   * undefined-valued leftovers dropped — exactly what the old copy produced,
   * since it skipped `style`/`comment`/`address` and any `undefined` value.
   */
  _setCompactFromModel(col: number, address: string, cellModel: RowModelCell): void {
    const type = cellModel.type as number;
    if (!isCompactLoadType(type)) {
      const cell = this.getCellEx({
        address,
        row: this._number,
        col,
      });
      cell.model = cellModel as never;
      return;
    }

    const value = valueFromCellModel(cellModel);
    const style = cellModel.style as Partial<Style> & Record<string, unknown>;
    const comment = cellModel.comment as {type?: string} | undefined;
    delete cellModel.style;
    delete cellModel.comment;
    // writeModelFromCellModel skipped undefined values — drop them so the
    // reused object has exactly the same shape (e.g. reconciled `styleId`,
    // hyperlink juggling leaves `styleId`/`result`/`value` undefined).
    for (const key in cellModel) {
      if (cellModel[key] === undefined) delete cellModel[key];
    }
    cellModel.address = address;

    const compact: CompactCell = {
      _c: 1,
      col,
      address,
      type,
      value: value as CellValue,
      wm: cellModel,
    };

    if (style) {
      compact.style = style;
    }

    if (comment) {
      if (comment.type === 'note') {
        compact.comment = Note.fromModel(comment as Parameters<typeof Note.fromModel>[0]);
      }
    }

    this._cells[col - 1] = compact;
  }

  findCell(colNumber: number): Cell | undefined {
    const entry = this._cells[colNumber - 1];
    if (!entry) return undefined;
    if (isCompact(entry)) {
      return this._materialize(colNumber, entry);
    }
    return entry;
  }

  // given {address, row, col}, find or create new cell
  getCellEx(address: RowCellAddress): Cell {
    const entry = this._cells[address.col - 1];
    if (entry) {
      if (isCompact(entry)) {
        return this._materialize(address.col, entry);
      }
      return entry;
    }
    const column = this._worksheet.getColumn(address.col);
    const cell = new Cell(this, column, address.address, {validateAddress: false});
    this._cells[address.col - 1] = cell;
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
    const entry = this._cells[col - 1];
    if (entry) {
      if (isCompact(entry)) {
        return this._materialize(col, entry);
      }
      return entry;
    }
    return this.getCellEx({
      address: colCache.encodeAddress(this._number, col),
      row: this._number,
      col,
    });
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
        cDst = this.findCell(i);
        cSrc = this.findCell(i - nExpand);
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
        cSrc = this.findCell(i);
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
      this._cells.forEach((entry, index) => {
        if (!entry) return;
        if (isCompact(entry)) {
          // type -1 = JSON object (has a value); Null is empty
          if (entry.type === Enums.ValueType.Null) return;
          fn!(this._materialize(index + 1, entry), index + 1);
          return;
        }
        if (entry.type !== Enums.ValueType.Null) {
          fn!(entry, index + 1);
        }
      });
    }
  }

  /**
   * Visit non-null cells without materializing CompactCell → Cell.
   * Used by doc-to-plain (and similar projections) so load stays lazy.
   */
  eachValue(callback: (info: CellValueInfo, colNumber: number) => void): void {
    const slots = this._cells;
    const n = slots.length;
    for (let i = 0; i < n; i++) {
      const entry = slots[i];
      if (!entry) continue;

      if (isCompact(entry)) {
        if (entry.type === Enums.ValueType.Null) continue;
        const note = entry.comment ? entry.comment.note : undefined;
        callback(
          {
            col: entry.col,
            address: entry.address,
            type: entry.type,
            value: entry.value,
            style: entry.style,
            note,
          },
          entry.col,
        );
        continue;
      }

      if (entry.type === Enums.ValueType.Null) continue;
      callback(
        {
          col: entry.col,
          address: entry.address,
          type: entry.type,
          value: entry.value,
          style: entry.style,
          note: entry.note,
        },
        entry.col,
      );
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
    this._cells.forEach(entry => {
      if (!entry) return;
      if (isCompact(entry)) {
        if (entry.type !== Enums.ValueType.Null) {
          values[entry.col] = entry.value as CellValue;
        }
        return;
      }
      if (entry.type !== Enums.ValueType.Null) {
        values[entry.col] = entry.value;
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
          this._setCompact(index + offset, item);
        }
      });
    } else {
      // assume object with column keys
      const obj = value as Record<string, CellValue>;
      this._worksheet.eachColumnKey((column, key) => {
        if (obj[key] !== undefined) {
          this._setCompact(column.number, obj[key], column);
        }
      });
    }
  }

  // returns true if the row includes at least one cell with a value
  get hasValues(): boolean {
    return this._cells.some((entry: CellSlot | undefined) => {
      if (!entry) return false;
      if (isCompact(entry)) {
        return entry.type !== Enums.ValueType.Null;
      }
      return entry.type !== Enums.ValueType.Null;
    });
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
    this._cells.forEach(entry => {
      if (!entry) return;
      const type = isCompact(entry) ? entry.type : entry.type;
      const col = isCompact(entry) ? entry.col : entry.col;
      if (type !== Enums.ValueType.Null) {
        if (!min || min > col) {
          min = col;
        }
        if (max < col) {
          max = col;
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
    this._cells.forEach(entry => {
      if (!entry) return;
      if (isCompact(entry)) {
        if (!entry.style) entry.style = {};
        (entry.style as Record<string, unknown>)[name] = value;
      } else {
        (entry as unknown as Record<string, unknown>)[name] = value;
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
    const cells: RowModelCell[] = [];
    let min = 0;
    let max = 0;
    const slots = this._cells;
    const n = slots.length;

    for (let i = 0; i < n; i++) {
      const entry = slots[i];
      if (!entry) continue;

      if (isCompact(entry)) {
        if (!min || min > entry.col) min = entry.col;
        if (max < entry.col) max = entry.col;
        cells.push(compactToModel(entry));
        continue;
      }

      const cellModel = entry.model;
      if (cellModel) {
        if (!min || min > entry.col) min = entry.col;
        if (max < entry.col) max = entry.col;
        cells.push(cellModel as RowModelCell);
      }
    }

    return this.height || cells.length
      ? {
          cells,
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
          // Prefer CompactCell on load — avoid Cell/Value strategy for common types
          this._setCompactFromModel(address!.col, address!.address, cellModel);
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
