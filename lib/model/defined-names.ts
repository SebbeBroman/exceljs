import colCache from '../utils/col-cache.js';
import CellMatrix from '../utils/cell-matrix.js';
import Range from './range.js';
import type {DefinedNamesModel} from './schema.js';
import type {MatrixCell} from '../utils/cell-matrix.js';
import type {CellAddress, DecodeExResult} from '../utils/col-cache.js';

// re-export MatrixCell shape used with mark flag
export type DefinedNameCell = MatrixCell & {mark?: boolean};

export interface DefinedNameLocation {
  sheetName?: string;
  top?: number;
  left?: number;
  bottom?: number;
  right?: number;
  row?: number;
  col?: number;
  address?: string;
  $col$row?: string;
}

export interface DefinedNameRanges {
  name: string;
  ranges: string[];
}

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

class DefinedNames {
  matrixMap: Record<string, CellMatrix>;

  constructor() {
    this.matrixMap = {};
  }

  getMatrix(name: string): CellMatrix {
    const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
    return matrix;
  }

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    this.addEx(location as DefinedNameLocation, name);
  }

  addEx(location: DefinedNameLocation | DecodeExResult, name: string): void {
    const matrix = this.getMatrix(name);
    const loc = location as DefinedNameLocation;
    if (loc.top) {
      for (let col = loc.left!; col <= loc.right!; col++) {
        for (let row = loc.top; row <= loc.bottom!; row++) {
          const address = {
            sheetName: loc.sheetName ?? '',
            address: colCache.n2l(col) + row,
            row,
            col,
            $col$row: `$${colCache.n2l(col)}$${row}`,
          };

          matrix.addCellEx(address as DecodeExResult);
        }
      }
    } else {
      matrix.addCellEx(location as DecodeExResult);
    }
  }

  remove(locStr: string, name: string): void {
    const location = colCache.decodeEx(locStr);
    this.removeEx(location as DefinedNameLocation, name);
  }

  removeEx(location: DefinedNameLocation | DecodeExResult, name: string): void {
    const matrix = this.getMatrix(name);
    matrix.removeCellEx(location as CellAddress);
  }

  removeAllNames(location: DefinedNameLocation | DecodeExResult): void {
    for (const matrix of Object.values(this.matrixMap)) {
      matrix.removeCellEx(location as CellAddress);
    }
  }

  forEach(callback: (name: string, cell: DefinedNameCell) => void): void {
    for (const [name, matrix] of Object.entries(this.matrixMap)) {
      matrix.forEach((cell: DefinedNameCell) => {
        callback(name, cell);
      });
    }
  }

  // get all the names of a cell
  getNames(addressStr: string): string[] {
    return this.getNamesEx(colCache.decodeEx(addressStr) as DefinedNameLocation);
  }

  getNamesEx(address: DefinedNameLocation | DecodeExResult): string[] {
    return Object.entries(this.matrixMap)
      .map(([name, matrix]) => matrix.findCellEx(address as DecodeExResult) && name)
      .filter(Boolean) as string[];
  }

  _explore(matrix: CellMatrix, cell: DefinedNameCell): Range {
    cell.mark = false;
    const {sheetName} = cell;

    const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    let x: number;
    let y: number;

    // grow vertical - only one col to worry about
    function vGrow(yy: number, edge: 'top' | 'bottom'): boolean {
      const c = matrix.findCellAt(sheetName, yy, cell.col) as DefinedNameCell | null | undefined;
      if (!c || !c.mark) {
        return false;
      }
      range[edge] = yy;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, 'top'); y--);
    for (y = cell.row + 1; vGrow(y, 'bottom'); y++);

    // grow horizontal - ensure all rows can grow
    function hGrow(xx: number, edge: 'left' | 'right'): boolean {
      const cells: DefinedNameCell[] = [];
      for (y = range.top; y <= range.bottom; y++) {
        const c = matrix.findCellAt(sheetName, y, xx) as DefinedNameCell | null | undefined;
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = xx;
      for (let i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--);
    for (x = cell.col + 1; hGrow(x, 'right'); x++);

    return range;
  }

  getRanges(name: string, matrix?: CellMatrix | null): DefinedNameRanges {
    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      return {name, ranges: []};
    }

    // mark and sweep!
    matrix.forEach((cell: DefinedNameCell) => {
      cell.mark = true;
    });
    const ranges = matrix
      .map((cell: DefinedNameCell) => (cell.mark ? this._explore(matrix!, cell) : false))
      .filter((r): r is Range => Boolean(r))
      .map((range: Range) => range.$shortRange);

    return {
      name,
      ranges,
    };
  }

  normaliseMatrix(matrix: CellMatrix, sheetName: string): void {
    // some of the cells might have shifted on specified sheet
    // need to reassign rows, cols
    matrix.forEachInSheet(
      sheetName,
      (cell: DefinedNameCell | null | undefined, row: number, col: number) => {
        if (cell) {
          if (cell.row !== row || cell.col !== col) {
            cell.row = row;
            cell.col = col;
            cell.address = colCache.n2l(col) + row;
          }
        }
      },
    );
  }

  spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    for (const matrix of Object.values(this.matrixMap)) {
      matrix.spliceRows(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    }
  }

  spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
    for (const matrix of Object.values(this.matrixMap)) {
      matrix.spliceColumns(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    }
  }

  get model(): DefinedNamesModel {
    // To get names per cell - just iterate over all names finding cells if they exist
    return Object.entries(this.matrixMap)
      .map(([name, matrix]) => this.getRanges(name, matrix))
      .filter((definedName: DefinedNameRanges) => definedName.ranges.length) as DefinedNamesModel;
  }

  set model(value: DefinedNamesModel) {
    // value is [ { name, ranges }, ... ]
    const matrixMap: Record<string, CellMatrix> = (this.matrixMap = {});
    value.forEach(definedName => {
      const matrix = (matrixMap[definedName.name] = new CellMatrix());
      definedName.ranges.forEach((rangeStr: string) => {
        if (rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
}

export default DefinedNames;
export {DefinedNames};
