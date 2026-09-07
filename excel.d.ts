/**
 * Public types for @sebbebroman/excel-ts 5.x (builder-first).
 *
 * Package `"types"` and `exports["."].types` / `exports["./node"].types` resolve here.
 * Legacy ExcelJS class typings remain in `index.d.ts` for internal shapes and
 * shared Style/Font/PageSetup definitions only — not the supported public API.
 */

import type {
  Alignment,
  Borders,
  CellErrorValue,
  CellFormulaValue,
  CellHyperlinkValue,
  CellRichTextValue,
  CellSharedFormulaValue,
  Comment,
  ConditionalFormattingOptions,
  DataValidation,
  Fill,
  Font,
  HeaderFooter,
  PageSetup,
  Protection,
  Style as FullStyle,
  TableProperties,
  WorkbookProperties,
  WorkbookView as LegacyWorkbookView,
  WorksheetProtection,
  WorksheetView,
} from './index.js';

export type Style = Partial<
  Pick<FullStyle, 'numFmt'> & {
    font: Partial<Font>;
    alignment: Partial<Alignment>;
    protection: Partial<Protection>;
    border: Partial<Borders>;
    fill: Fill;
  }
>;

export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | undefined
  | CellErrorValue
  | CellRichTextValue
  | CellHyperlinkValue
  | CellFormulaValue
  | CellSharedFormulaValue;

export interface WorkbookMeta {
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
  company?: string;
  manager?: string;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  language?: string;
  revision?: Date | string | number;
  contentStatus?: string;
  properties?: Partial<WorkbookProperties>;
  views?: LegacyWorkbookView[];
}

export interface ColumnInput {
  /**
   * Header label. Only the first line is used when `header` is an array
   * (multi-row headers are not supported — extra lines are dropped).
   */
  header?: string | string[];
  key?: string;
  width?: number;
  hidden?: boolean;
  style?: Style;
  outlineLevel?: number;
}

export type RowInput = ReadonlyArray<CellValue> | Record<string, CellValue>;

export type SheetTitleInput = string | {
  text: string;
  /** Applied to `merge` when given, otherwise to the title cell. */
  style?: Style;
  /** Valid A1 range (e.g. `'A1:B1'`). Invalid ranges throw at build time. */
  merge?: string;
};

export interface SheetCell {
  value: CellValue;
  style?: Style;
}

export interface SheetRow {
  number: number;
  cells: Record<number, SheetCell>;
  height?: number;
  hidden?: boolean;
  style?: Style;
}

export type NoteValue = string | Comment;
export type WorksheetViewInput = Partial<WorksheetView>;
export type ProtectOptions = Partial<WorksheetProtection> & {spinCount?: number};

export interface MediaImage {
  extension: 'jpeg' | 'png' | 'gif';
  base64?: string;
  filename?: string;
  buffer?: unknown;
}

export type SheetImageRange =
  | string
  | {
      tl: {col: number; row: number};
      br: {col: number; row: number};
      editAs?: string;
      hyperlinks?: {hyperlink: string; tooltip?: string};
    }
  | {
      tl: {col: number; row: number};
      ext: {width: number; height: number};
      editAs?: string;
      hyperlinks?: {hyperlink: string; tooltip?: string};
    };

export interface ProtectConfig {
  password?: string;
  options?: ProtectOptions;
}

export interface DefinedNameEntry {
  name: string;
  refersTo: string;
}

export interface SheetImagePlacement {
  imageId: number;
  range: SheetImageRange;
}

export interface SheetModel {
  id: number;
  name: string;
  rows: SheetRow[];
  columns?: ColumnInput[];
  merges?: string[];
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  dataValidations?: Record<string, DataValidation>;
  conditionalFormattings?: ConditionalFormattingOptions[];
  notes?: Record<string, NoteValue>;
  protect?: ProtectConfig;
  sheetProtection?: Record<string, unknown>;
  tables?: TableProperties[];
  images?: SheetImagePlacement[];
}

export interface Workbook {
  meta: WorkbookMeta;
  sheets: SheetModel[];
  media?: MediaImage[];
  definedNames?: DefinedNameEntry[];
}

export interface WorkbookInit extends WorkbookMeta {
  sheets?: SheetModel[];
  media?: MediaImage[];
  definedNames?: DefinedNameEntry[];
}

export interface WriteOptions {
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: unknown;
}

export interface LoadOptions {
  /** Treat input as base64 when data is a string. */
  base64?: boolean;
  /** XML node names to ignore while parsing. */
  ignoreNodes?: string[];
}

export type ReadRowsFormat = 'auto' | 'csv' | 'xlsx';
export type ViewFormat = ReadRowsFormat;

export type ColSlice =
  | {start?: number; end?: number}
  | number[]
  | string[];

export interface ViewWorkbookOptions {
  format?: ViewFormat;
  /** Optional name for format sniffing (e.g. `file.name`). */
  filename?: string;
  /** @deprecated Prefer `filename`. */
  name?: string;
  encoding?: string;
}

export interface RowsOptions {
  /** 1-based inclusive start row. Default `1`. */
  start?: number;
  /** 1-based inclusive end row. Default last used row. */
  end?: number;
  cols?: ColSlice;
  range?: string;
  values?: 'string' | 'cell';
  blankrows?: boolean;
  trim?: boolean;
  defval?: string;
}

export interface RecordsOptions extends RowsOptions {
  header?: boolean | number;
}

export interface SheetView {
  readonly name: string;
  readonly id: number;
  readonly index: number;
  readonly protected: boolean;
  rows(options?: RowsOptions & {values?: 'string'}): string[][];
  rows(options: RowsOptions & {values: 'cell'}): CellValue[][];
  rows(options?: RowsOptions): string[][] | CellValue[][];
  records(options?: RecordsOptions): Record<string, string | CellValue>[];
}

export interface WorkbookView {
  readonly format: 'csv' | 'xlsx';
  readonly sheetNames: string[];
  readonly meta: WorkbookMeta;
  sheet(nameOrIndex?: string | number): SheetView;
  toJSON(): Workbook;
}

/** Options for `readRows` (CSV/xlsx → dense `string[][]`). */
export interface ReadRowsOptions extends ViewWorkbookOptions, RowsOptions {
  /** Sheet name or 0-based index. Default `0`. */
  sheet?: string | number;
}

export interface SheetInit {
  title?: SheetTitleInput;
  columns?: ColumnInput[];
  rows?: RowInput[];
  merges?: string[];
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
}

/** Options for `csv.parse` / Node `readCsvFile`. */
export interface CsvParseOptions {
  /** Sheet name used by Node `readCsvFile` (default `"Sheet1"`). Ignored by `csv.parse`. */
  sheetName?: string;
  dateFormats?: string[];
  map?: (datum: string, index?: number) => unknown;
  /** Pass-through to fast-csv parse options. */
  parserOptions?: Record<string, unknown>;
}

/** Options for `csv.stringify` / builder `.csv()` / Node `writeCsvFile`. */
export interface CsvStringifyOptions {
  /** Sheet name to export (default: active/first sheet). */
  sheetName?: string;
  /** Sheet model `id` or 1-based index into `sheets`. */
  sheetId?: number;
  dateFormat?: string;
  dateUTC?: boolean;
  map?: (value: unknown, index?: number) => unknown;
  /** When true (default), emit blank lines for missing row numbers. */
  includeEmptyRows?: boolean;
  /** Pass-through to fast-csv format options. */
  formatterOptions?: Record<string, unknown>;
  /** File encoding for Node `writeCsvFile` only. */
  encoding?: BufferEncoding;
}

export interface SheetBuilder {
  title(title: SheetTitleInput): SheetBuilder;
  row(values: RowInput): SheetBuilder;
  rows(values: RowInput[]): SheetBuilder;
  cell(address: string, value: CellValue, style?: Style): SheetBuilder;
  cells(map: Record<string, CellValue>): SheetBuilder;
  style(range: string, style: Style): SheetBuilder;
  merge(range: string): SheetBuilder;
  columns(cols: ColumnInput[]): SheetBuilder;
  views(views: WorksheetViewInput[]): SheetBuilder;
  pageSetup(setup: Partial<PageSetup>): SheetBuilder;
  headerFooter(hf: Partial<HeaderFooter>): SheetBuilder;
  dataValidation(address: string, rules: DataValidation): SheetBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): SheetBuilder;
  note(address: string, note: NoteValue): SheetBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): SheetBuilder;
  table(table: TableProperties): SheetBuilder;
  image(imageId: number, range: SheetImageRange): SheetBuilder;
}

export interface WorkbookBuilder {
  sheet(name: string, init?: SheetInit | ((s: SheetBuilder) => void)): WorkbookBuilder;
  /** Emit a title row on the active sheet (call `.sheet(name)` first). */
  title(title: SheetTitleInput): WorkbookBuilder;
  row(values: RowInput): WorkbookBuilder;
  rows(values: RowInput[]): WorkbookBuilder;
  cell(address: string, value: CellValue, style?: Style): WorkbookBuilder;
  cells(map: Record<string, CellValue>): WorkbookBuilder;
  style(range: string, style: Style): WorkbookBuilder;
  merge(range: string): WorkbookBuilder;
  columns(cols: ColumnInput[]): WorkbookBuilder;
  props(meta: WorkbookInit): WorkbookBuilder;
  views(views: WorksheetViewInput[]): WorkbookBuilder;
  pageSetup(setup: Partial<PageSetup>): WorkbookBuilder;
  headerFooter(hf: Partial<HeaderFooter>): WorkbookBuilder;
  dataValidation(address: string, rules: DataValidation): WorkbookBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): WorkbookBuilder;
  note(address: string, note: NoteValue): WorkbookBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): WorkbookBuilder;
  table(table: TableProperties): WorkbookBuilder;
  /** Register workbook media; returns image id. */
  image(def: MediaImage): number;
  /** Alias for `image(def)` — clearer name for the register step. */
  addImage(def: MediaImage): number;
  /** Place a registered image on the active sheet. */
  image(imageId: number, range: SheetImageRange): WorkbookBuilder;
  definedName(name: string, refersTo: string): WorkbookBuilder;
  build(): Workbook;
  writeBuffer(opts?: WriteOptions): Promise<Uint8Array>;
  /** Stringify the active (or first) sheet as CSV. */
  csv(opts?: CsvStringifyOptions): Promise<string>;
}

/** Named CSV helpers. */
export const csv: {
  parse(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
  stringify(input: Workbook | WorkbookBuilder, opts?: CsvStringifyOptions): Promise<string>;
};

export function parseCsv(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
export function stringifyCsv(
  input: Workbook | WorkbookBuilder,
  opts?: CsvStringifyOptions,
): Promise<string>;

export function workbook(
  init?: WorkbookInit | Workbook | WorkbookView,
): WorkbookBuilder;
export function isWorkbookBuilder(value: unknown): value is WorkbookBuilder;
export function writeBuffer(
  input: Workbook | WorkbookBuilder,
  options?: WriteOptions,
): Promise<Uint8Array>;
export function load(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
  options?: LoadOptions,
): Promise<Workbook>;

/** Read-only view over CSV or OOXML (xlsx/xlsm/…). */
export function viewWorkbook(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ViewWorkbookOptions,
): Promise<WorkbookView>;
export function isWorkbookView(value: unknown): value is WorkbookView;

/**
 * CSV or xlsx → dense `string[][]`.
 * Sugar: `(await viewWorkbook(data, opts)).sheet(opts.sheet ?? 0).rows({ values: 'string', ... })`
 */
export function readRows(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ReadRowsOptions,
): Promise<string[][]>;
export function cellToDisplayString(value: CellValue, defval?: string): string;

export const ValueType: typeof import('./index.js').ValueType;
export const FormulaType: typeof import('./index.js').FormulaType;
export const RelationshipType: typeof import('./index.js').RelationshipType;
export const DocumentType: typeof import('./index.js').DocumentType;
export const ReadingOrder: typeof import('./index.js').ReadingOrder;
export const ErrorValue: typeof import('./index.js').ErrorValue;
export const enums: unknown;

/** Node entry (`@sebbebroman/excel-ts/node`) also exports (Node-only — do not import from browser bundles): */
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
// Re-export feature types used by builder methods
export type {
  DataValidation,
  ConditionalFormattingOptions,
  PageSetup,
  HeaderFooter,
  TableProperties,
  Comment,
  WorksheetView,
  WorksheetProtection,
};
