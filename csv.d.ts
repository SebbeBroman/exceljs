import type {
  SheetInit,
  Workbook,
  WorkbookBuilder,
  WorkbookView,
  ViewWorkbookOptions,
  ReadRowsOptions,
} from './excel.js';
/** Options for `csv.parse` / Node `readCsvFile`. */
export interface CsvParseOptions {
  /** Sheet name used by Node `readCsvFile` (default `"Sheet1"`). Ignored by `csv.parse`. */
  sheetName?: string;
  dateFormats?: string[];
  map?: (datum: string, index?: number) => unknown;
  /** Pass-through to fast-csv parse options. */
  parserOptions?: Record<string, unknown>;
}

/** Options for `csv.stringify` / Node `writeCsvFile`. */
export interface CsvStringifyOptions {
  /** Sheet name to export (default: first sheet). */
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

export const csv: {
  parse(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
  stringify(input: Workbook | WorkbookBuilder, opts?: CsvStringifyOptions): Promise<string>;
};

export function parseCsv(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
export function stringifyCsv(
  input: Workbook | WorkbookBuilder,
  opts?: CsvStringifyOptions,
): Promise<string>;

/**
 * Create a builder, optionally seeded from a `WorkbookInit`, a plain `Workbook`
 * snapshot, or a read-only `WorkbookView`.
 *
 * WARNING: `WorkbookView` (`viewWorkbook`/`readRows`) is values-only (no styles,
 * merges, formulas, hyperlinks, images; dates arrive numeric) — `workbook(view)`
 * re-encodes from values, so never use it for fidelity round-trips. Use `load()`
 * when formatting must survive.
 */
export function viewCsv(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ViewWorkbookOptions,
): Promise<WorkbookView>;
export function readCsvRows(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ReadRowsOptions,
): Promise<string[][]>;
