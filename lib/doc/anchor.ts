import colCache from '../utils/col-cache.js';
import type {Worksheet} from '../../index.js';

/** Address-like input accepted by Anchor constructor. */
export interface AnchorAddress {
  nativeCol?: number;
  nativeColOff?: number;
  nativeRow?: number;
  nativeRowOff?: number;
  col?: number;
  row?: number;
}

export interface AnchorModel {
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;
}

/** Minimal worksheet surface used by Anchor for col/row dimensions. */
export interface AnchorWorksheet {
  getColumn(col: number): {isCustomWidth?: boolean; width?: number} | undefined;
  getRow(row: number): {height?: number} | undefined;
}

class Anchor {
  worksheet: AnchorWorksheet | Worksheet | null | undefined;
  nativeCol!: number;
  nativeColOff!: number;
  nativeRow!: number;
  nativeRowOff!: number;

  constructor(
    worksheet?: AnchorWorksheet | Worksheet | null,
    address?: string | AnchorAddress | null,
    offset = 0,
  ) {
    this.worksheet = worksheet;

    if (!address) {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    } else if (typeof address === 'string') {
      const decoded = colCache.decodeAddress(address);
      this.nativeCol = (decoded.col ?? 0) + offset;
      this.nativeColOff = 0;
      this.nativeRow = (decoded.row ?? 0) + offset;
      this.nativeRowOff = 0;
    } else if (address.nativeCol !== undefined) {
      this.nativeCol = address.nativeCol || 0;
      this.nativeColOff = address.nativeColOff || 0;
      this.nativeRow = address.nativeRow || 0;
      this.nativeRowOff = address.nativeRowOff || 0;
    } else if (address.col !== undefined) {
      this.col = address.col + offset;
      this.row = (address.row ?? 0) + offset;
    } else {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    }
  }

  static asInstance(model: Anchor | AnchorAddress | null | undefined): Anchor | null | undefined {
    return model instanceof Anchor || model == null ? model : new Anchor(model as unknown as AnchorWorksheet);
  }

  get col(): number {
    return this.nativeCol + Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth;
  }

  set col(v: number) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  }

  get row(): number {
    return this.nativeRow + Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight;
  }

  set row(v: number) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  }

  get colWidth(): number {
    return this.worksheet &&
      this.worksheet.getColumn(this.nativeCol + 1) &&
      this.worksheet.getColumn(this.nativeCol + 1)!.isCustomWidth
      ? Math.floor((this.worksheet.getColumn(this.nativeCol + 1)!.width ?? 0) * 10000)
      : 640000;
  }

  get rowHeight(): number {
    return this.worksheet &&
      this.worksheet.getRow(this.nativeRow + 1) &&
      this.worksheet.getRow(this.nativeRow + 1)!.height
      ? Math.floor((this.worksheet.getRow(this.nativeRow + 1)!.height ?? 0) * 10000)
      : 180000;
  }

  get model(): AnchorModel {
    return {
      nativeCol: this.nativeCol,
      nativeColOff: this.nativeColOff,
      nativeRow: this.nativeRow,
      nativeRowOff: this.nativeRowOff,
    };
  }

  set model(value: AnchorModel) {
    this.nativeCol = value.nativeCol;
    this.nativeColOff = value.nativeColOff;
    this.nativeRow = value.nativeRow;
    this.nativeRowOff = value.nativeRowOff;
  }
}

export default Anchor;
export {Anchor};
