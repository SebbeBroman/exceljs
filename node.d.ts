import type {CsvParseOptions, CsvStringifyOptions} from './csv.js';
/**
 * Node-only types for `@sebbebroman/exceljs/node`.
 *
 * Re-exports the full browser-safe API from the main entry plus Node-only
 * file/stream helpers. Do not import this entry from browser bundles
 * (it pulls `node:fs`).
 */
export * from '@sebbebroman/exceljs';

import type {
  Workbook,
  WorkbookBuilder,
  WriteOptions,
  LoadOptions,
  ColumnInput,
  RowInput,
  WorksheetViewInput,
  PageSetup,
  HeaderFooter,
} from '@sebbebroman/exceljs';

/** Node entry (`@sebbebroman/exceljs/node`) also exports (Node-only — do not import from browser bundles): */
export function writeFile(
  path: string,
  input: Workbook | WorkbookBuilder,
  options?: WriteOptions,
): Promise<void>;
export function readFile(path: string, options?: LoadOptions): Promise<Workbook>;
export function readCsvFile(path: string, options?: CsvParseOptions): Promise<Workbook>;
export function writeCsvFile(
  path: string,
  input: Workbook | WorkbookBuilder,
  options?: CsvStringifyOptions,
): Promise<void>;

// --- Streaming (Node-only: `./node`) ---

export interface StreamWriteOptions {
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: unknown;
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
}

export interface StreamSheetOptions {
  columns?: ColumnInput[];
  state?: string;
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  autoFilter?: unknown;
  properties?: unknown;
}

export interface StreamWriteSheetSpec extends StreamSheetOptions {
  name: string;
  rows: AsyncIterable<RowInput> | Iterable<RowInput>;
}

export interface StreamWriteDeclarative extends StreamWriteOptions {
  sheets: StreamWriteSheetSpec[];
}

export interface StreamSheetHandle {
  columns(cols: ColumnInput[]): StreamSheetHandle;
  row(values: RowInput): StreamSheetHandle;
  rows(values: AsyncIterable<RowInput> | Iterable<RowInput>): Promise<void>;
}

export interface StreamWorkbookHandle {
  sheet(name: string, options?: StreamSheetOptions): StreamSheetHandle;
}

export type StreamWriteCallback = (w: StreamWorkbookHandle) => void | Promise<void>;
export type StreamWriteSpec = StreamWriteDeclarative | StreamWriteCallback;

/**
 * Stream an xlsx workbook to a path or Node Writable.
 * Rows are committed as written (bounded memory). Node-only.
 *
 * @example Declarative
 * ```ts
 * await streamWrite('out.xlsx', {
 *   sheets: [{ name: 'Data', rows: bigIterable }],
 * });
 * ```
 *
 * @example Callback
 * ```ts
 * await streamWrite('out.xlsx', async w => {
 *   const s = w.sheet('Data', { columns: [{ header: 'Id', key: 'id' }] });
 *   for await (const row of source) s.row(row);
 * });
 * ```
 */
export function streamWrite(
  dest: string | import('node:stream').Writable,
  spec: StreamWriteSpec,
  options?: StreamWriteOptions,
): Promise<void>;

export interface StreamReadOptions {
  worksheets?: 'emit' | 'ignore' | string;
  sharedStrings?: 'cache' | 'emit' | 'ignore' | string;
  hyperlinks?: 'cache' | 'emit' | 'ignore' | string;
  styles?: 'cache' | 'ignore' | string;
  entries?: 'emit' | 'ignore' | string;
}

export interface StreamReadRow {
  sheetName: string;
  sheetId: number | string;
  rowNumber: number;
  /** Sparse: index 0 unused; column A is `values[1]`. */
  values: unknown[];
}

/**
 * Async-iterate rows from an xlsx path or readable stream (Node-only).
 * Prefer `readFile` / `load` for small files and full model access.
 */
export function streamRead(
  input: string | import('node:stream').Readable,
  options?: StreamReadOptions,
): AsyncGenerator<StreamReadRow>;
