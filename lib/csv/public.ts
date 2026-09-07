/**
 * Named CSV API (no side-effect registration).
 *
 *   import { csv, workbook } from '@sebbebroman/excel-ts';
 *
 *   const init = await csv.parse('a,b\n1,2');
 *   const text = await workbook().sheet('Data', init).csv();
 *   // or: await csv.stringify(workbook().sheet('Data', init));
 */

import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat.js';
import utc from 'dayjs/plugin/utc.js';
import fastCsv from 'fast-csv';
import type {CellValue, RowInput, SheetInit, SheetModel, Workbook} from '../model/types.js';
import type {WorkbookBuilder} from '../builder/workbook-builder.js';

// dayjs default export is callable + has .extend; typings vary by version
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const dayjsFn = dayjs as any;
// Lazily extend once on first use (keeps module import side-effect free so
// `sideEffects: false` tree-shaking stays correct for shared-dayjs consumers).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let dayjsExt: any | null = null;
function ext(): any {
  if (!dayjsExt) dayjsExt = dayjsFn.extend(customParseFormat).extend(utc);
  return dayjsExt;
}

const DEFAULT_DATE_FORMATS = [
  'YYYY-MM-DD[T]HH:mm:ssZ',
  'YYYY-MM-DD[T]HH:mm:ss',
  'MM-DD-YYYY',
  'YYYY-MM-DD',
];

const SpecialValues: Record<string, boolean | {error: string}> = {
  true: true,
  false: false,
  '#N/A': {error: '#N/A'},
  '#REF!': {error: '#REF!'},
  '#NAME?': {error: '#NAME?'},
  '#DIV/0!': {error: '#DIV/0!'},
  '#NULL!': {error: '#NULL!'},
  '#VALUE!': {error: '#VALUE!'},
  '#NUM!': {error: '#NUM!'},
};

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

function defaultParseMap(dateFormats: string[]) {
  return function mapDatum(datum: string): unknown {
    if (datum === '') {
      return null;
    }
    const datumNumber = Number(datum);
    if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
      return datumNumber;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dt = dateFormats.reduce<any>((matchingDate, currentDateFormat) => {
      if (matchingDate) {
        return matchingDate;
      }
      const dayjsObj = ext()(datum, currentDateFormat, true);
      if (dayjsObj.isValid()) {
        return dayjsObj;
      }
      return null;
    }, null);
    if (dt) {
      return new Date(dt.valueOf());
    }
    const special = SpecialValues[datum];
    if (special !== undefined) {
      return special;
    }
    return datum;
  };
}

function defaultStringifyMap(options: CsvStringifyOptions) {
  const {dateFormat, dateUTC} = options;
  return (value: unknown): unknown => {
    if (value) {
      const v = value as Record<string, unknown>;
      if (v.text || v.hyperlink) {
        return (v.hyperlink as string) || (v.text as string) || '';
      }
      if (v.formula || v.result !== undefined) {
        return v.result ?? '';
      }
      if (value instanceof Date) {
        if (dateFormat) {
          return dateUTC ? ext().utc(value).format(dateFormat) : ext()(value).format(dateFormat);
        }
        return dateUTC ? ext().utc(value).format() : ext()(value).format();
      }
      if (v.error) {
        return v.error;
      }
      if (typeof value === 'object') {
        return JSON.stringify(value);
      }
    }
    return value;
  };
}

function isBuilderLike(value: unknown): value is WorkbookBuilder {
  return Boolean(
    value &&
      typeof value === 'object' &&
      '_ops' in value &&
      Array.isArray((value as WorkbookBuilder)._ops) &&
      typeof (value as WorkbookBuilder).build === 'function',
  );
}

function toPlainWorkbook(input: Workbook | WorkbookBuilder): Workbook {
  if (isBuilderLike(input)) {
    return input.build();
  }
  return input;
}

function resolveSheet(wb: Workbook, options?: CsvStringifyOptions): SheetModel {
  if (options?.sheetName) {
    const found = wb.sheets.find(s => s.name === options.sheetName);
    if (!found) {
      throw new Error(`Sheet not found: ${options.sheetName}`);
    }
    return found;
  }
  if (options?.sheetId != null) {
    const byId = wb.sheets.find(s => s.id === options.sheetId);
    if (byId) return byId;
    const byIndex = wb.sheets[options.sheetId - 1];
    if (byIndex) return byIndex;
    throw new Error(`Sheet not found: ${options.sheetId}`);
  }
  const first = wb.sheets[0];
  if (!first) {
    throw new Error('Workbook has no sheets');
  }
  return first;
}

function sheetToRowArrays(sheet: SheetModel, options: CsvStringifyOptions): unknown[][] {
  const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
  const map = options.map || defaultStringifyMap(options);
  const sorted = [...sheet.rows].sort((a, b) => a.number - b.number);
  const out: unknown[][] = [];
  let lastRow = 1;

  for (const row of sorted) {
    if (includeEmptyRows) {
      while (lastRow++ < row.number - 1) {
        out.push([]);
      }
    }

    const colNums = Object.keys(row.cells).map(Number);
    const maxCol = colNums.length ? Math.max(...colNums) : 0;
    const values: unknown[] = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.cells[c];
      values.push(map(cell?.value, c - 1));
    }
    out.push(values);
    lastRow = row.number;
  }

  return out;
}

/**
 * Parse CSV text into a `SheetInit` suitable for `workbook().sheet(name, init)`.
 */
export async function parseCsv(text: string, opts?: CsvParseOptions): Promise<SheetInit> {
  const options = opts || {};
  const dateFormats = options.dateFormats || DEFAULT_DATE_FORMATS;
  const map = options.map || defaultParseMap(dateFormats);

  const rows: RowInput[] = await new Promise((resolve, reject) => {
    const collected: RowInput[] = [];
    const stream = fastCsv.parseString(text, options.parserOptions as Parameters<typeof fastCsv.parseString>[1]);
    stream
      .on('data', (data: string[] | Record<string, string>) => {
        if (Array.isArray(data)) {
          collected.push(data.map((datum, i) => map(datum, i) as CellValue));
        } else {
          // headers: true → object rows; keep as keyed RowInput
          const obj: Record<string, CellValue> = {};
          for (const [key, datum] of Object.entries(data)) {
            obj[key] = map(String(datum ?? ''), undefined) as CellValue;
          }
          collected.push(obj);
        }
      })
      .on('end', () => resolve(collected))
      .on('error', reject);
  });

  return {rows};
}

/**
 * Stringify a plain Workbook or builder to CSV text (one sheet).
 */
export async function stringifyCsv(
  input: Workbook | WorkbookBuilder,
  opts?: CsvStringifyOptions,
): Promise<string> {
  const options = opts || {};
  const wb = toPlainWorkbook(input);
  const sheet = resolveSheet(wb, options);
  const rows = sheetToRowArrays(sheet, options);
  return fastCsv.writeToString(rows as string[][], options.formatterOptions as Parameters<
    typeof fastCsv.writeToString
  >[1]);
}

/** Named CSV helpers. */
export const csv = {
  parse: parseCsv,
  stringify: stringifyCsv,
};

export default csv;
