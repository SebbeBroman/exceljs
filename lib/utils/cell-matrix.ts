import colCache from './col-cache.js';
import type {CellAddress, DecodeExResult, RangeAddressEx} from './col-cache.js';

export interface MatrixCell {
  sheetName: string;
  address: string;
  row: number;
  col: number;
  [key: string]: unknown;
}

type SheetRow = Array<MatrixCell | null | undefined>;
type Sheet = Array<SheetRow | undefined>;

class CellMatrix {
  template: unknown;
  sheets: Record<string, Sheet>;

  constructor(template?: unknown) {
    this.template = template;
    this.sheets = {};
  }

  addCell(addressStr: string): void {
    this.addCellEx(colCache.decodeEx(addressStr));
  }

  getCell(addressStr: string): MatrixCell | undefined {
    return this.findCellEx(colCache.decodeEx(addressStr), true);
  }

  findCell(addressStr: string): MatrixCell | undefined {
    return this.findCellEx(colCache.decodeEx(addressStr), false);
  }

  findCellAt(
    sheetName: string,
    rowNumber: number,
    colNumber: number,
  ): MatrixCell | null | undefined {
    const sheet = this.sheets[sheetName];
    const row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  }

  addCellEx(address: DecodeExResult): void {
    if ('top' in address && address.top) {
      const range = address as RangeAddressEx;
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) {
          this.getCellAt(range.sheetName as string, row, col);
        }
      }
    } else {
      this.findCellEx(address, true);
    }
  }

  getCellEx(address: DecodeExResult): MatrixCell | undefined {
    return this.findCellEx(address, true);
  }

  findCellEx(address: DecodeExResult, create?: boolean): MatrixCell | undefined {
    const sheet = this.findSheet(address, create);
    const row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  }

  getCellAt(sheetName: string, rowNumber: number, colNumber: number): MatrixCell {
    const sheet = this.sheets[sheetName] || (this.sheets[sheetName] = []);
    const row = sheet[rowNumber] || (sheet[rowNumber] = []);
    const cell =
      row[colNumber] ||
      (row[colNumber] = {
        sheetName,
        address: colCache.n2l(colNumber) + rowNumber,
        row: rowNumber,
        col: colNumber,
      });
    return cell;
  }

  removeCellEx(address: CellAddress): void {
    const sheet = this.findSheet(address);
    if (!sheet) {
      return;
    }
    const row = this.findSheetRow(sheet, address);
    if (!row) {
      return;
    }
    delete row[address.col as number];
  }

  forEachInSheet(
    sheetName: string,
    callback: (cell: MatrixCell, rowNumber: number, colNumber: number) => void,
  ): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      sheet.forEach((row, rowNumber) => {
        if (row) {
          row.forEach((cell, colNumber) => {
            if (cell) {
              callback(cell, rowNumber, colNumber);
            }
          });
        }
      });
    }
  }

  forEach(
    callback: (cell: MatrixCell, rowNumber: number, colNumber: number) => void,
  ): void {
    for (const sheetName of Object.keys(this.sheets)) {
      this.forEachInSheet(sheetName, callback);
    }
  }

  map<T>(callback: (cell: MatrixCell) => T): T[] {
    const results: T[] = [];
    this.forEach(cell => {
      results.push(callback(cell));
    });
    return results;
  }

  findSheet(address: DecodeExResult, create?: boolean): Sheet | undefined {
    const name = (address as CellAddress).sheetName as string;
    if (this.sheets[name]) {
      return this.sheets[name];
    }
    if (create) {
      return (this.sheets[name] = []);
    }
    return undefined;
  }

  findSheetRow(
    sheet: Sheet | undefined,
    address: DecodeExResult,
    create?: boolean,
  ): SheetRow | undefined {
    const {row} = address as CellAddress;
    if (sheet && sheet[row as number]) {
      return sheet[row as number];
    }
    if (create) {
      return (sheet![row as number] = []);
    }
    return undefined;
  }

  findRowCell(
    row: SheetRow | undefined,
    address: DecodeExResult,
    create?: boolean,
  ): MatrixCell | undefined {
    const {col} = address as CellAddress;
    if (row && row[col as number]) {
      return row[col as number] as MatrixCell;
    }
    if (create) {
      return (row![col as number] = this.template
        ? (Object.assign(address, JSON.parse(JSON.stringify(this.template))) as MatrixCell)
        : (address as unknown as MatrixCell));
    }
    return undefined;
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      const inserts: SheetRow[] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push([]);
      }
      sheet.splice(start, numDelete, ...inserts);
    }
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets[sheetName];
    if (sheet) {
      const inserts: null[] = [];
      for (let i = 0; i < numInsert; i++) {
        inserts.push(null);
      }
      for (const row of sheet) {
        if (row) {
          row.splice(start, numDelete, ...inserts);
        }
      }
    }
  }
}

export default CellMatrix;
export {CellMatrix};
