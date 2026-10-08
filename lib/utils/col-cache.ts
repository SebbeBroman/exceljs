const addressRegex = /^[A-Z]+\d+$/;

export interface CellAddress {
  address: string;
  col: number | undefined;
  row: number | undefined;
  $col$row: string;
  sheetName?: string;
}

export interface RangeAddress {
  top: number;
  left: number;
  bottom: number;
  right: number;
  tl: string;
  br: string;
  dimensions: string;
}

export interface RangeAddressEx {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;
  tl: CellAddress;
  br: CellAddress;
  dimensions: string;
}

export interface ErrorAddress {
  sheetName?: string;
  error: string;
}

export type DecodeResult = CellAddress | RangeAddress;
export type DecodeExResult = CellAddress | RangeAddressEx | ErrorAddress;

// =========================================================================
// Column Letter to Number conversion
// Memoize common columns only; a far-right lookup must not allocate all 16,384.
const columnLetters: string[] = [];
const addresses: Record<string, CellAddress> = Object.create(null);
const addressKeys: string[] = [];
let nextAddressKey = 0;
const ADDRESS_CACHE_LIMIT = 2048;
function cacheAddress(key: string, address: CellAddress): void {
  if (!addresses[key]) {
    if (addressKeys.length < ADDRESS_CACHE_LIMIT) addressKeys.push(key);
    else {
      delete addresses[addressKeys[nextAddressKey]];
      addressKeys[nextAddressKey] = key;
      nextAddressKey = (nextAddressKey + 1) % ADDRESS_CACHE_LIMIT;
    }
  }
  addresses[key] = address;
}

const colCache = {
  l2n(letter: string): number {
    if (typeof letter !== 'string' || !/^[A-Z]{1,3}$/.test(letter)) {
      throw new Error(`Out of bounds. Invalid column letter: ${letter}`);
    }
    let number = 0;
    for (let i = 0; i < letter.length; i++) number = number * 26 + letter.charCodeAt(i) - 64;
    if (number > 16384) throw new Error(`Out of bounds. Invalid column letter: ${letter}`);
    return number;
  },
  n2l(number: number): string {
    if (!Number.isInteger(number) || number < 1 || number > 16384) {
      throw new Error(`${number} is out of bounds. Excel supports columns from 1 to 16384`);
    }
    const cached = number <= 256 && columnLetters[number];
    if (cached) return cached;
    let value = number;
    let letter = '';
    while (value > 0) {
      value--;
      letter = String.fromCharCode(65 + (value % 26)) + letter;
      value = Math.floor(value / 26);
    }
    if (number <= 256) columnLetters[number] = letter;
    return letter;
  },

  // =========================================================================
  // Address processing

  // check if value looks like an address
  validateAddress(value: string): true {
    if (!addressRegex.test(value)) {
      throw new Error(`Invalid Address: ${value}`);
    }
    return true;
  },

  // convert address string into structure
  decodeAddress(value: string): CellAddress {
    const addr = value.length <= 8 && addresses[value];
    if (addr) {
      return addr;
    }
    let hasCol = false;
    let col = '';
    let colNumber: number | undefined = 0;
    let hasRow = false;
    let row = '';
    let rowNumber: number | undefined = 0;
    for (let i = 0, char: number; i < value.length; i++) {
      char = value.charCodeAt(i);
      // col should before row
      if (!hasRow && char >= 65 && char <= 90) {
        // 65 = 'A'.charCodeAt(0)
        // 90 = 'Z'.charCodeAt(0)
        hasCol = true;
        col += value[i];
        // colNumber starts from 1
        colNumber = (colNumber as number) * 26 + char - 64;
      } else if (char >= 48 && char <= 57) {
        // 48 = '0'.charCodeAt(0)
        // 57 = '9'.charCodeAt(0)
        hasRow = true;
        row += value[i];
        // rowNumber starts from 0
        rowNumber = (rowNumber as number) * 10 + char - 48;
      } else if (hasRow && hasCol && char !== 36) {
        // 36 = '$'.charCodeAt(0)
        break;
      }
    }
    if (!hasCol) {
      colNumber = undefined;
    } else if ((colNumber as number) > 16384) {
      throw new Error(`Out of bounds. Invalid column letter: ${col}`);
    }
    if (!hasRow) {
      rowNumber = undefined;
    }

    // in case $row$col
    value = col + row;

    const address: CellAddress = {
      address: value,
      col: colNumber,
      row: rowNumber,
      $col$row: `$${col}$${row}`,
    };

    // Keep hot small addresses bounded rather than retaining the whole 100x100 square.
    if (
      colNumber !== undefined &&
      rowNumber !== undefined &&
      colNumber <= 100 &&
      rowNumber <= 100
    ) {
      cacheAddress(value, address);
      cacheAddress(address.$col$row, address);
    }

    return address;
  },

  // convert r,c into structure (if only 1 arg, assume r is address string)
  getAddress(r: string | number, c?: number): CellAddress {
    if (c) {
      const address = this.n2l(c) + r;
      return this.decodeAddress(address);
    }
    return this.decodeAddress(r as string);
  },

  // convert [address], [tl:br] into address structures
  decode(value: string): DecodeResult {
    const parts = value.split(':');
    if (parts.length === 2) {
      const tl = this.decodeAddress(parts[0]);
      const br = this.decodeAddress(parts[1]);
      const result: RangeAddress = {
        top: Math.min(tl.row as number, br.row as number),
        left: Math.min(tl.col as number, br.col as number),
        bottom: Math.max(tl.row as number, br.row as number),
        right: Math.max(tl.col as number, br.col as number),
        tl: '',
        br: '',
        dimensions: '',
      };
      // reconstruct tl, br and dimensions
      result.tl = this.n2l(result.left) + result.top;
      result.br = this.n2l(result.right) + result.bottom;
      result.dimensions = `${result.tl}:${result.br}`;
      return result;
    }
    return this.decodeAddress(value);
  },

  // convert [sheetName!][$]col[$]row[[$]col[$]row] into address or range structures
  decodeEx(value: string): DecodeExResult {
    const groups = value.match(/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/)!;

    const sheetName = groups[1] || groups[2]; // Qouted and unqouted groups
    const reference = groups[3]; // Remaining address

    const parts = reference.split(':');
    if (parts.length > 1) {
      const tl = this.decodeAddress(parts[0]);
      const br = this.decodeAddress(parts[1]);
      const top = Math.min(tl.row as number, br.row as number);
      const left = Math.min(tl.col as number, br.col as number);
      const bottom = Math.max(tl.row as number, br.row as number);
      const right = Math.max(tl.col as number, br.col as number);

      const leftLetters = Number.isFinite(left) ? this.n2l(left) : '';
      const rightLetters = Number.isFinite(right) ? this.n2l(right) : '';
      const tlStr = leftLetters + (Number.isFinite(top) ? top : '');
      const brStr = rightLetters + (Number.isFinite(bottom) ? bottom : '');

      return {
        top,
        left,
        bottom,
        right,
        sheetName,
        tl: {
          address: tlStr,
          col: left,
          row: top,
          $col$row: `$${leftLetters}$${Number.isFinite(top) ? top : ''}`,
          sheetName,
        },
        br: {
          address: brStr,
          col: right,
          row: bottom,
          $col$row: `$${rightLetters}$${Number.isFinite(bottom) ? bottom : ''}`,
          sheetName,
        },
        dimensions: `${tlStr}:${brStr}`,
      };
    }
    if (reference.startsWith('#')) {
      return sheetName ? {sheetName, error: reference} : {error: reference};
    }

    const address = this.decodeAddress(reference);
    return sheetName ? {sheetName, ...address} : address;
  },

  // convert row,col into address string
  encodeAddress(row: number, col: number): string {
    return colCache.n2l(col) + row;
  },

  // convert row,col into string address or t,l,b,r into range
  encode(...args: number[]): string {
    switch (args.length) {
      case 2:
        return colCache.encodeAddress(args[0], args[1]);
      case 4:
        return `${colCache.encodeAddress(args[0], args[1])}:${colCache.encodeAddress(
          args[2],
          args[3],
        )}`;
      default:
        throw new Error('Can only encode with 2 or 4 arguments');
    }
  },

  // return true if address is contained within range
  inRange(range: number[], address: number[]): boolean {
    const [left, top, , right, bottom] = range;
    const [col, row] = address;
    return col >= left && col <= right && row >= top && row <= bottom;
  },
};

export default colCache;
export {colCache};
