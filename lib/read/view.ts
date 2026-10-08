/**
 * Read-only workbook view.
 *
 *   import { viewWorkbook } from '@sebbebroman/exceljs';
 *   import { workbook, writeBuffer } from '@sebbebroman/exceljs'; // write (separate)
 *
 *   const view = await viewWorkbook(data, { format: 'auto', filename: file.name });
 *   const rows = view.sheet(0).rows({ start: 1, end: 100, cols: { end: 4 } });
 *   const wb = workbook(view);
 */

import {parseText} from '@sebbebroman/fast-csv';
import type {CellValue, SheetModel, Workbook, WorkbookMeta} from '../model/types.js';
import colCache from '../utils/col-cache.js';
import {cellToDisplayString} from './cells.js';
import {decodeText, sniffFormat, toUint8Array, type ViewFormat} from './format.js';
import {parseA1Range, resolveSlice, sheetExtent, type ColSlice} from './slice.js';
import {
  openLightPackage,
  parseLightSheet,
  type LightPackage,
  type LightParseSheetOptions,
  type LightSheetGrid,
  type LightSheetInfo,
} from './xlsx-light.js';

/** Brand for `workbook(view)` / `isWorkbookView`. */
export const WORKBOOK_VIEW = Symbol.for('@sebbebroman/exceljs.WorkbookView');

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
  return parseText(text, {headers: false, ignoreEmpty: false, trim: false});
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

/** Convert light cell grid (absolute row numbers via blankrows) into SheetModel. */
function lightGridToSheetModel(grid: LightSheetGrid, startRow = 1): SheetModel {
  const rows: SheetModel['rows'] = [];
  for (let r = 0; r < grid.rows.length; r++) {
    const line = grid.rows[r] as CellValue[];
    const cells: Record<number, {value: CellValue}> = {};
    for (let c = 0; c < line.length; c++) {
      const v = line[c];
      if (v != null && v !== '') {
        cells[c + 1] = {value: v};
      }
    }
    rows.push({number: startRow + r, cells});
  }
  return {id: grid.id, name: grid.name, rows};
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

function headerKeys(headerLine: (string | CellValue)[], asString: boolean): string[] {
  return headerLine.map((h, i) => {
    const raw = asString ? String(h ?? '') : cellToDisplayString(h as CellValue, '');
    const s = raw.trim();
    return s || `col${i + 1}`;
  });
}

/** Map view RowsOptions → light parse options (including A1 range). */
export function rowsOptionsToLight(options: RowsOptions): LightParseSheetOptions {
  let start = options.start;
  let end = options.end;
  let cols: LightParseSheetOptions['cols'] | undefined;

  if (options.range) {
    const r = parseA1Range(options.range);
    start = start ?? r.rowStart;
    end = end ?? r.rowEnd;
    if (options.cols == null) {
      cols = {start: r.colStart, end: r.colEnd};
    }
  }

  if (cols == null && options.cols != null) {
    const raw = options.cols;
    if (Array.isArray(raw)) {
      if (raw.length === 0) {
        cols = [];
      } else if (typeof raw[0] === 'number') {
        cols = raw as number[];
      } else {
        cols = (raw as string[]).map(letter => {
          const s = letter.trim().toUpperCase();
          if (/^\d+$/.test(s)) return Number(s);
          return colCache.l2n(s);
        });
      }
    } else {
      cols = raw;
    }
  }

  return {
    start,
    end,
    cols,
    values: options.values ?? 'string',
    blankrows: options.blankrows,
    trim: options.trim,
    defval: options.defval,
  };
}

// ---------------------------------------------------------------------------
// CSV sheet view (eager model)
// ---------------------------------------------------------------------------

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
    return recordsFromMatrix(this.model, options ?? {});
  }
}

function recordsFromMatrix(
  model: SheetModel,
  opts: RecordsOptions,
): Record<string, string | CellValue>[] {
  const asString = opts.values !== 'cell';
  const headerOpt = opts.header === undefined ? true : opts.header;

  if (headerOpt === false) {
    const data = buildRowMatrix(model, opts);
    return data.matrix.map(line => {
      const rec: Record<string, string | CellValue> = {};
      line.forEach((v, i) => {
        rec[`col${i + 1}`] = v;
      });
      return rec;
    });
  }

  if (typeof headerOpt === 'number') {
    const headerMatrix = buildRowMatrix(model, {
      ...opts,
      start: headerOpt,
      end: headerOpt,
      blankrows: true,
    });
    const keys = headerKeys(headerMatrix.matrix[0] ?? [], asString);
    const bodyStart = Math.max(opts.start ?? 1, headerOpt + 1);
    const body = buildRowMatrix(model, {...opts, start: bodyStart});
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
  const headerMatrix = buildRowMatrix(model, {
    ...opts,
    start: sliceStart,
    end: sliceStart,
    blankrows: true,
  });
  if (!headerMatrix.matrix.length) return [];
  const keys = headerKeys(headerMatrix.matrix[0]!, asString);
  const body = buildRowMatrix(model, {
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

// ---------------------------------------------------------------------------
// Lazy xlsx view (per-sheet parse)
// ---------------------------------------------------------------------------

class LazyXlsxSheetView implements SheetView {
  constructor(
    private readonly pkg: LightPackage,
    private readonly ref: LightSheetInfo,
    private readonly cache: Map<number, SheetModel>,
  ) {}

  get name(): string {
    return this.ref.name;
  }

  get id(): number {
    return this.ref.id;
  }

  get index(): number {
    return this.ref.index;
  }

  /** Light path does not parse sheetProtection. */
  get protected(): boolean {
    return false;
  }

  /** Full sheet → SheetModel, cached after first materialize. */
  private materialize(): SheetModel {
    let model = this.cache.get(this.ref.index);
    if (!model) {
      const grid = parseLightSheet(this.pkg, this.ref.index, {
        values: 'cell',
        blankrows: true,
      });
      model = lightGridToSheetModel(grid, 1);
      this.cache.set(this.ref.index, model);
    }
    return model;
  }

  rows(options?: RowsOptions & {values?: 'string'}): string[][];
  rows(options: RowsOptions & {values: 'cell'}): CellValue[][];
  rows(options?: RowsOptions): string[][] | CellValue[][] {
    const opts = options ?? {};
    const cached = this.cache.get(this.ref.index);
    if (cached) {
      const {matrix, asString} = buildRowMatrix(cached, opts);
      return (asString ? matrix : matrix) as string[][] | CellValue[][];
    }

    // Lazy: parse only this sheet's XML with slice options (supports end early-stop).
    const lightOpts = rowsOptionsToLight(opts);
    const grid = parseLightSheet(this.pkg, this.ref.index, lightOpts);
    const asString = opts.values !== 'cell';
    if (asString) return grid.rows as string[][];
    return grid.rows as CellValue[][];
  }

  records(options?: RecordsOptions): Record<string, string | CellValue>[] {
    // records() issues multiple slices (header + body) — materialize once.
    return recordsFromMatrix(this.materialize(), options ?? {});
  }
}

class LazyXlsxWorkbookView implements WorkbookView {
  readonly [WORKBOOK_VIEW] = true as const;
  readonly format = 'xlsx' as const;
  readonly meta: WorkbookMeta = {};
  /** Per-sheet SheetModel after first full materialize. */
  private readonly sheetCache = new Map<number, SheetModel>();

  constructor(private readonly pkg: LightPackage) {}

  get sheetNames(): string[] {
    return this.pkg.sheetNames;
  }

  sheet(nameOrIndex: string | number = 0): SheetView {
    let ref: LightSheetInfo | undefined;
    if (typeof nameOrIndex === 'number') {
      ref = this.pkg.sheets[nameOrIndex];
    } else {
      ref = this.pkg.sheets.find(s => s.name === nameOrIndex);
    }
    if (!ref) {
      throw new Error(`Sheet not found: ${String(nameOrIndex)}`);
    }
    return new LazyXlsxSheetView(this.pkg, ref, this.sheetCache);
  }

  toJSON(): Workbook {
    const sheets: SheetModel[] = [];
    for (const ref of this.pkg.sheets) {
      let model = this.sheetCache.get(ref.index);
      if (!model) {
        const grid = parseLightSheet(this.pkg, ref.index, {
          values: 'cell',
          blankrows: true,
        });
        model = lightGridToSheetModel(grid, 1);
        this.sheetCache.set(ref.index, model);
      }
      sheets.push(model);
    }
    return {meta: this.meta, sheets};
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
    typeof data === 'string'
      ? data
      : toUint8Array(data as ArrayBuffer | Uint8Array | ArrayBufferView);

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
  const pkg = await openLightPackage(bytes);
  return new LazyXlsxWorkbookView(pkg);
}
