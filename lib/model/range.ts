import colCache from '../utils/col-cache.js';
import type {Location} from './schema.js';
import type {CellAddress, DecodeExResult, RangeAddressEx} from '../utils/col-cache.js';

export interface RangeModel {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;
}

/** Minimal row surface for expandRow. */
export interface RangeRowLike {
  dimensions: {min: number; max: number} | null;
  number: number;
}

// used by worksheet to calculate sheet dimensions
class Range {
  model!: RangeModel;

  // Overload-friendly constructor matching original argument forms
  constructor(...args: unknown[]) {
    this.decode(args);
  }

  setTLBR(
    t: number | string,
    l: number | string,
    b?: number | string,
    r?: number,
    s?: string,
  ): void {
    if (arguments.length < 4) {
      // setTLBR(tl, br, s)
      const tl = colCache.decodeAddress(t as string);
      const br = colCache.decodeAddress(l as string);
      this.model = {
        top: Math.min(tl.row ?? 0, br.row ?? 0),
        left: Math.min(tl.col ?? 0, br.col ?? 0),
        bottom: Math.max(tl.row ?? 0, br.row ?? 0),
        right: Math.max(tl.col ?? 0, br.col ?? 0),
        sheetName: b as string | undefined,
      };

      this.setTLBR(tl.row ?? 0, tl.col ?? 0, br.row ?? 0, br.col ?? 0, s);
    } else {
      // setTLBR(t, l, b, r, s)
      const top = t as number;
      const left = l as number;
      const bottom = b as number;
      const right = r as number;
      this.model = {
        top: Math.min(top, bottom),
        left: Math.min(left, right),
        bottom: Math.max(top, bottom),
        right: Math.max(left, right),
        sheetName: s,
      };
    }
  }

  decode(argv: IArguments | unknown[]): void {
    const args = argv as unknown[];
    switch (args.length) {
      case 5: // [t,l,b,r,s]
        this.setTLBR(
          args[0] as number,
          args[1] as number,
          args[2] as number,
          args[3] as number,
          args[4] as string,
        );
        break;
      case 4: // [t,l,b,r]
        this.setTLBR(args[0] as number, args[1] as number, args[2] as number, args[3] as number);
        break;

      case 3: // [tl,br,s]
        this.setTLBR(args[0] as string, args[1] as string, args[2] as string);
        break;
      case 2: // [tl,br]
        this.setTLBR(args[0] as string, args[1] as string);
        break;

      case 1: {
        const value = args[0] as Range | unknown[] | RangeModel | string;
        if (value instanceof Range) {
          // copy constructor
          this.model = {
            top: value.model.top,
            left: value.model.left,
            bottom: value.model.bottom,
            right: value.model.right,
            sheetName: value.sheetName,
          };
        } else if (value instanceof Array) {
          // an arguments array
          this.decode(value);
        } else if (
          value &&
          typeof value === 'object' &&
          (value as RangeModel).top &&
          (value as RangeModel).left &&
          (value as RangeModel).bottom &&
          (value as RangeModel).right
        ) {
          // a model
          const m = value as RangeModel;
          this.model = {
            top: m.top,
            left: m.left,
            bottom: m.bottom,
            right: m.right,
            sheetName: m.sheetName,
          };
        } else {
          // [sheetName!]tl:br
          const tlbr = colCache.decodeEx(value as string) as DecodeExResult;
          if ('top' in tlbr && tlbr.top) {
            const range = tlbr as RangeAddressEx;
            this.model = {
              top: range.top,
              left: range.left,
              bottom: range.bottom,
              right: range.right,
              sheetName: range.sheetName,
            };
          } else {
            const cell = tlbr as CellAddress & {row?: number; col?: number; sheetName?: string};
            this.model = {
              top: cell.row ?? 0,
              left: cell.col ?? 0,
              bottom: cell.row ?? 0,
              right: cell.col ?? 0,
              sheetName: cell.sheetName,
            };
          }
        }
        break;
      }

      case 0:
        this.model = {
          top: 0,
          left: 0,
          bottom: 0,
          right: 0,
        };
        break;

      default:
        throw new Error(`Invalid number of arguments to _getDimensions() - ${args.length}`);
    }
  }

  get top(): number {
    return this.model.top || 1;
  }

  set top(value: number) {
    this.model.top = value;
  }

  get left(): number {
    return this.model.left || 1;
  }

  set left(value: number) {
    this.model.left = value;
  }

  get bottom(): number {
    return this.model.bottom || 1;
  }

  set bottom(value: number) {
    this.model.bottom = value;
  }

  get right(): number {
    return this.model.right || 1;
  }

  set right(value: number) {
    this.model.right = value;
  }

  get sheetName(): string | undefined {
    return this.model.sheetName;
  }

  set sheetName(value: string | undefined) {
    this.model.sheetName = value;
  }

  get _serialisedSheetName(): string {
    const {sheetName} = this.model;
    if (sheetName) {
      if (/^[a-zA-Z0-9]*$/.test(sheetName)) {
        return `${sheetName}!`;
      }
      return `'${sheetName}'!`;
    }
    return '';
  }

  expand(top: number, left: number, bottom: number, right: number): void {
    if (!this.model.top || top < this.top) this.top = top;
    if (!this.model.left || left < this.left) this.left = left;
    if (!this.model.bottom || bottom > this.bottom) this.bottom = bottom;
    if (!this.model.right || right > this.right) this.right = right;
  }

  expandRow(row: RangeRowLike | null | undefined): void {
    if (row) {
      const {dimensions, number} = row;
      if (dimensions) {
        this.expand(number, dimensions.min, number, dimensions.max);
      }
    }
  }

  expandToAddress(addressStr: string): void {
    const address = colCache.decodeEx(addressStr) as CellAddress & {row: number; col: number};
    this.expand(address.row, address.col, address.row, address.col);
  }

  get tl(): string {
    return colCache.n2l(this.left) + this.top;
  }

  get $t$l(): string {
    return `$${colCache.n2l(this.left)}$${this.top}`;
  }

  get br(): string {
    return colCache.n2l(this.right) + this.bottom;
  }

  get $b$r(): string {
    return `$${colCache.n2l(this.right)}$${this.bottom}`;
  }

  get range(): string {
    return `${this._serialisedSheetName + this.tl}:${this.br}`;
  }

  get $range(): string {
    return `${this._serialisedSheetName + this.$t$l}:${this.$b$r}`;
  }

  get shortRange(): string {
    return this.count > 1 ? this.range : this._serialisedSheetName + this.tl;
  }

  get $shortRange(): string {
    return this.count > 1 ? this.$range : this._serialisedSheetName + this.$t$l;
  }

  get count(): number {
    return (1 + this.bottom - this.top) * (1 + this.right - this.left);
  }

  toString(): string {
    return this.range;
  }

  intersects(other: Range | (Location & {sheetName?: string})): boolean {
    if (other.sheetName && this.sheetName && other.sheetName !== this.sheetName) return false;
    if (other.bottom < this.top) return false;
    if (other.top > this.bottom) return false;
    if (other.right < this.left) return false;
    if (other.left > this.right) return false;
    return true;
  }

  contains(addressStr: string): boolean {
    const address = colCache.decodeEx(addressStr);
    return this.containsEx(address as Partial<{sheetName: string; row: number; col: number}>);
  }

  containsEx(address: Partial<{sheetName: string; row: number; col: number}>): boolean {
    if (address.sheetName && this.sheetName && address.sheetName !== this.sheetName) return false;
    return (
      (address.row ?? 0) >= this.top &&
      (address.row ?? 0) <= this.bottom &&
      (address.col ?? 0) >= this.left &&
      (address.col ?? 0) <= this.right
    );
  }

  forEachAddress(cb: (address: string, row: number, col: number) => void): void {
    for (let col = this.left; col <= this.right; col++) {
      for (let row = this.top; row <= this.bottom; row++) {
        cb(colCache.encodeAddress(row, col), row, col);
      }
    }
  }
}

export default Range;
export {Range};
