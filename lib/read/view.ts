/**
 * Read-only workbook view.
 *
 *   import { viewWorkbook } from '@sebbebroman/excel-ts';
 *   import { workbook, writeBuffer } from '@sebbebroman/excel-ts'; // write (separate)
 *
 *   const view = await viewWorkbook(data, { format: 'auto', filename: file.name });
 *   const rows = view.sheet(0).rows({ start: 1, end: 100, cols: { end: 4 } });
 *   const wb = workbook(view);
 */

import fastCsv from 'fast-csv';
import {load} from '../xlsx/load.js';
import type {CellValue, SheetModel, Workbook, WorkbookMeta} from '../model/types.js';
import {cellToDisplayString} from './cells.js';
import {decodeText, sniffFormat, toUint8Array, type ViewFormat} from './format.js';
import {resolveSlice, sheetExtent, type ColSlice} from './slice.js';

/** Brand for `workbook(view)` / `isWorkbookView`. */
export const WORKBOOK_VIEW = Symbol.for('@sebbebroman/excel-ts.WorkbookView');

export interface ViewWorkbookOptions {
  /** `'auto'` (default), `'csv'`, or OOXML `'xlsx'` (also .xlsm etc.). */
  format?: ViewFormat;
  /** Optional name for format sniffing (e.g. `file.name`). */
  filename?: string;
  /** @deprecated Prefer `filename`. */
  name?: string;
  /** CSV TextDecoder encoding when data is binary. Default `'utf-8'`. */
  encoding?: string;
}

export interface RowsOptions {
  /** 1-based inclusive start row. Default `1`. */
  start?: number;
  /** 1-based inclusive end row. Default last used row. */
  end?: number;
  /**
   * Column slice:
   * - `{ start, end }` contiguous 1-based
   * - `number[]` / `string[]` discrete columns (`['A','C']` → A and C only)
   * - span via `range: 'A1:D10'` or `{ start: 1, end: 4 }`
   */
  cols?: ColSlice;
  /** A1 range (e.g. `'B2:F100'`); row start/end can still narrow. */
  range?: string;
  /** `'string'` (default) → `string[][]`; `'cell'` → `CellValue[][]`. */
  values?: 'string' | 'cell';
  /**
   * When `false` (default), skip entirely empty rows in the result.
   * When `true`, keep blank rows inside the slice.
   */
  blankrows?: boolean;
  /** Trim string cells. Default `true` when `values: 'string'`. */
  trim?: boolean;
  /** Empty cell placeholder for string mode. Default `''`. */
  defval?: string;
}

export interface RecordsOptions extends RowsOptions {
  /**
   * `true` (default) = first row of the **slice** is the header.
   * `number` = absolute 1-based sheet row used as header keys.
   * `false` = keys `col1`, `col2`, …
   */
  header?: boolean | number;
}

export interface SheetView {
  readonly name: string;
  readonly id: number;
  readonly index: number;
  /** True if the sheet has protection metadata (cell values are still readable). */
  readonly protected: boolean;

  rows(options?: RowsOptions & {values?: 'string'}): string[][];
  rows(options: RowsOptions & {values: 'cell'}): CellValue[][];
  rows(options?: RowsOptions): string[][] | CellValue[][];

  records(options?: RecordsOptions): Record<string, string | CellValue>[];
}

export interface WorkbookView {
  readonly [WORKBOOK_VIEW]: true;
  readonly format: 'csv' | 'xlsx';
  readonly sheetNames: string[];
  readonly meta: WorkbookMeta;

  /** Sheet by **0-based** index or name. */
  sheet(nameOrIndex?: string | number): SheetView;

  /** Plain snapshot for `workbook(view)` / serialization. */
  toJSON(): Workbook;
}

export function isWorkbookView(value: unknown): value is WorkbookView {
  return Boolean(
    value &&
      typeof value === 'object' &&
      (value as WorkbookView)[WORKBOOK_VIEW] === true &&
      typeof (value as WorkbookView).toJSON === 'function',
  );
}

async function parseCsvGrid(text: string): Promise<string[][]> {
  return new Promise((resolve, reject) => {
    const rows: string[][] = [];
    fastCsv
      .parseString(text, {headers: false, ignoreEmpty: false, trim: false})
      .on('error', reject)
      .on('data', (row: string[] | Record<string, string>) => {
        if (Array.isArray(row)) {
          rows.push(row.map(c => (c == null ? '' : String(c))));
        } else {
          rows.push(Object.values(row).map(c => (c == null ? '' : String(c))));
        }
      })
      .on('end', () => resolve(rows));
  });
}

function gridToSheetModel(grid: string[][], name: string, id: number): SheetModel {
  const rows: SheetModel['rows'] = [];
  for (let r = 0; r < grid.length; r++) {
    const line = grid[r]!;
    const cells: Record<number, {value: CellValue}> = {};
    for (let c = 0; c < line.length; c++) {
      cells[c + 1] = {value: line[c] ?? ''};
    }
    rows.push({number: r + 1, cells});
  }
  return {id, name, rows};
}

function buildRowMatrix(
  sheet: SheetModel,
  options: RowsOptions,
): {matrix: (string | CellValue)[][]; asString: boolean} {
  const asString = options.values !== 'cell';
  const trim = options.trim !== false && asString;
  const defval = options.defval ?? '';
  const skipBlank = options.blankrows !== true;

  const {maxRow, maxCol} = sheetExtent(sheet);
  if (maxRow === 0 || maxCol === 0) {
    return {matrix: [], asString};
  }

  const bounds = resolveSlice(options, maxRow, maxCol);
  const byNumber = new Map(sheet.rows.map(r => [r.number, r]));
  const matrix: (string | CellValue)[][] = [];

  for (let r = bounds.rowStart; r <= bounds.rowEnd; r++) {
    const src = byNumber.get(r);
    const line: (string | CellValue)[] = [];
    for (const c of bounds.cols) {
      const cell = src?.cells[c];
      if (asString) {
        let s = cell ? cellToDisplayString(cell.value, defval) : defval;
        if (trim) s = s.trim();
        line.push(s);
      } else {
        line.push(cell ? (cell.value ?? null) : null);
      }
    }
    if (skipBlank) {
      if (asString && line.every(v => v === '' || v === defval)) continue;
      if (!asString && line.every(v => v == null || v === '')) continue;
    }
    matrix.push(line);
  }

  return {matrix, asString};
}

function headerKeys(
  headerLine: (string | CellValue)[],
  asString: boolean,
): string[] {
  return headerLine.map((h, i) => {
    const raw = asString
      ? String(h ?? '')
      : cellToDisplayString(h as CellValue, '');
    const s = raw.trim();
    return s || `col${i + 1}`;
  });
}

class SheetViewImpl implements SheetView {
  constructor(
    private readonly model: SheetModel,
    readonly index: number,
  ) {}

  get name(): string {
    return this.model.name;
  }

  get id(): number {
    return this.model.id;
  }

  get protected(): boolean {
    return Boolean(this.model.sheetProtection);
  }

  rows(options?: RowsOptions & {values?: 'string'}): string[][];
  rows(options: RowsOptions & {values: 'cell'}): CellValue[][];
  rows(options?: RowsOptions): string[][] | CellValue[][] {
    const {matrix, asString} = buildRowMatrix(this.model, options ?? {});
    if (asString) return matrix as string[][];
    return matrix as CellValue[][];
  }

  records(options?: RecordsOptions): Record<string, string | CellValue>[] {
    const opts = options ?? {};
    const asString = opts.values !== 'cell';
    const headerOpt = opts.header === undefined ? true : opts.header;

    if (headerOpt === false) {
      const data = buildRowMatrix(this.model, opts);
      return data.matrix.map(line => {
        const rec: Record<string, string | CellValue> = {};
        line.forEach((v, i) => {
          rec[`col${i + 1}`] = v;
        });
        return rec;
      });
    }

    if (typeof headerOpt === 'number') {
      const headerMatrix = buildRowMatrix(this.model, {
        ...opts,
        start: headerOpt,
        end: headerOpt,
        blankrows: true,
      });
      const keys = headerKeys(headerMatrix.matrix[0] ?? [], asString);
      const bodyStart = Math.max(opts.start ?? 1, headerOpt + 1);
      const body = buildRowMatrix(this.model, {...opts, start: bodyStart});
      return body.matrix.map(line => {
        const rec: Record<string, string | CellValue> = {};
        keys.forEach((k, i) => {
          rec[k] = line[i] ?? (asString ? '' : null);
        });
        return rec;
      });
    }

    // header: true — first row of slice
    const sliceStart = opts.start ?? 1;
    const headerMatrix = buildRowMatrix(this.model, {
      ...opts,
      start: sliceStart,
      end: sliceStart,
      blankrows: true,
    });
    if (!headerMatrix.matrix.length) return [];
    const keys = headerKeys(headerMatrix.matrix[0]!, asString);
    const body = buildRowMatrix(this.model, {
      ...opts,
      start: sliceStart + 1,
    });
    return body.matrix.map(line => {
      const rec: Record<string, string | CellValue> = {};
      keys.forEach((k, i) => {
        rec[k] = line[i] ?? (asString ? '' : null);
      });
      return rec;
    });
  }
}

class WorkbookViewImpl implements WorkbookView {
  readonly [WORKBOOK_VIEW] = true as const;

  constructor(
    readonly format: 'csv' | 'xlsx',
    private readonly plain: Workbook,
  ) {}

  get sheetNames(): string[] {
    return this.plain.sheets.map(s => s.name);
  }

  get meta(): WorkbookMeta {
    return this.plain.meta ?? {};
  }

  sheet(nameOrIndex: string | number = 0): SheetView {
    let model: SheetModel | undefined;
    let index = 0;
    if (typeof nameOrIndex === 'number') {
      index = nameOrIndex;
      model = this.plain.sheets[nameOrIndex];
    } else {
      index = this.plain.sheets.findIndex(s => s.name === nameOrIndex);
      model = index >= 0 ? this.plain.sheets[index] : undefined;
    }
    if (!model) {
      throw new Error(`Sheet not found: ${String(nameOrIndex)}`);
    }
    return new SheetViewImpl(model, index);
  }

  toJSON(): Workbook {
    return this.plain;
  }
}

/**
 * Open CSV or OOXML bytes as a read-only {@link WorkbookView}.
 * Does not import the write/builder graph.
 */
export async function viewWorkbook(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ViewWorkbookOptions,
): Promise<WorkbookView> {
  const opts = options ?? {};
  const filename = opts.filename ?? opts.name;
  const binaryOrText =
    typeof data === 'string' ? data : toUint8Array(data as ArrayBuffer | Uint8Array | ArrayBufferView);

  const format = sniffFormat(
    typeof data === 'string' ? data : (binaryOrText as Uint8Array),
    filename,
    opts.format ?? 'auto',
  );

  if (format === 'csv') {
    const text = decodeText(
      typeof data === 'string' ? data : (binaryOrText as Uint8Array),
      opts.encoding ?? 'utf-8',
    );
    const grid = await parseCsvGrid(text);
    const sheet = gridToSheetModel(grid, 'Sheet1', 1);
    const plain: Workbook = {meta: {}, sheets: [sheet]};
    return new WorkbookViewImpl('csv', plain);
  }

  const bytes =
    typeof data === 'string' ? new TextEncoder().encode(data) : (binaryOrText as Uint8Array);
  const plain = await load(bytes);
  return new WorkbookViewImpl('xlsx', plain);
}
