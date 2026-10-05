import dayjs from 'dayjs';
import utils from '../utils/utils.js';
import fs from 'fs';
import * as fastCsv from '@sebbebroman/fast-csv/node';
import customParseFormat from 'dayjs/plugin/customParseFormat.js';
import utc from 'dayjs/plugin/utc.js';
import type {Readable, Writable} from 'node:stream';
import StreamBuf from '../utils/stream-buf.js';

// dayjs default export is callable + has .extend; typings vary by version
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const dayjsFn = dayjs as any;
const dayjsExt = dayjsFn.extend(customParseFormat).extend(utc);

const {
  fs: {exists},
} = utils as {fs: {exists: (filename: string) => Promise<boolean>}};

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

export interface CsvReadOptions {
  sheetName?: string;
  dateFormats?: string[];
  map?: (datum: string, index?: number) => unknown;
  parserOptions?: Record<string, unknown>;
}

export interface CsvWriteOptions {
  sheetName?: string;
  sheetId?: number;
  encoding?: string;
  dateFormat?: string;
  dateUTC?: boolean;
  map?: (value: unknown, index?: number) => unknown;
  includeEmptyRows?: boolean;
  formatterOptions?: Record<string, unknown>;
}

/** Minimal workbook surface used by CSV */
export interface CsvWorkbook {
  addWorksheet(name?: string): CsvWorksheet;
  getWorksheet(id?: string | number): CsvWorksheet | undefined;
}

export interface CsvWorksheet {
  addRow(values: unknown[]): unknown;
  eachRow(iteratee: (row: {values: unknown[]}, rowNumber: number) => void): void;
}

class CSV {
  workbook: CsvWorkbook;
  worksheet: CsvWorksheet | null;

  constructor(workbook: CsvWorkbook) {
    this.workbook = workbook;
    this.worksheet = null;
  }

  async readFile(filename: string, options?: CsvReadOptions): Promise<CsvWorksheet> {
    options = options || {};
    if (!(await exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    const worksheet = await this.read(stream, options);
    stream.close();
    return worksheet;
  }

  read(stream: Readable, options?: CsvReadOptions): Promise<CsvWorksheet> {
    options = options || {};

    return new Promise((resolve, reject) => {
      const worksheet = this.workbook.addWorksheet(options.sheetName);

      const dateFormats = options.dateFormats || [
        'YYYY-MM-DD[T]HH:mm:ssZ',
        'YYYY-MM-DD[T]HH:mm:ss',
        'MM-DD-YYYY',
        'YYYY-MM-DD',
      ];
      const map =
        options.map ||
        function (datum: string) {
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
            const dayjsObj = dayjsExt(datum, currentDateFormat, true);
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

      const csvStream = fastCsv
        .parse(options.parserOptions)
        .on('data', (data: string[]) => {
          worksheet.addRow(data.map(map));
        })
        .on('end', () => {
          csvStream.emit('worksheet', worksheet);
        });

      csvStream.on('worksheet', resolve).on('error', reject);

      stream.pipe(csvStream);
    });
  }

  /**
   * @deprecated since version 4.0. You should use `CSV#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
   */
  createInputStream(): never {
    throw new Error(
      '`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md',
    );
  }

  write(stream: Writable, options?: CsvWriteOptions): Promise<void> {
    return new Promise((resolve, reject) => {
      options = options || {};
      // const encoding = options.encoding || 'utf8';
      // const separator = options.separator || ',';
      // const quoteChar = options.quoteChar || '\'';

      const worksheet = this.workbook.getWorksheet(options.sheetName || options.sheetId);

      const csvStream = fastCsv.format(options.formatterOptions);
      stream.on('finish', () => {
        resolve();
      });
      csvStream.on('error', reject);
      csvStream.pipe(stream);

      const {dateFormat, dateUTC} = options;
      const map =
        options.map ||
        ((value: unknown) => {
          if (value) {
            const v = value as Record<string, unknown>;
            if (v.text || v.hyperlink) {
              return (v.hyperlink as string) || (v.text as string) || '';
            }
            if (v.formula || v.result) {
              return v.result || '';
            }
            if (value instanceof Date) {
              if (dateFormat) {
                return dateUTC
                  ? dayjsExt.utc(value).format(dateFormat)
                  : dayjsExt(value).format(dateFormat);
              }
              return dateUTC ? dayjsExt.utc(value).format() : dayjsExt(value).format();
            }
            if (v.error) {
              return v.error;
            }
            if (typeof value === 'object') {
              return JSON.stringify(value);
            }
          }
          return value;
        });

      const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
      let lastRow = 1;
      if (worksheet) {
        worksheet.eachRow((row, rowNumber) => {
          if (includeEmptyRows) {
            while (lastRow++ < rowNumber - 1) {
              csvStream.write([]);
            }
          }
          const {values} = row;
          // row.values is 1-indexed (index 0 unused)
          (values as unknown[]).shift();
          csvStream.write((values as unknown[]).map(map));
          lastRow = rowNumber;
        });
      }
      csvStream.end();
    });
  }

  writeFile(filename: string, options?: CsvWriteOptions): Promise<void> {
    options = options || {};

    const streamOptions = {
      encoding: (options.encoding || 'utf8') as BufferEncoding,
    };
    const stream = fs.createWriteStream(filename, streamOptions);

    return this.write(stream, options);
  }

  async writeBuffer(options?: CsvWriteOptions): Promise<unknown> {
    const stream = new StreamBuf();
    await this.write(stream as unknown as Writable, options);
    return stream.read();
  }
}

export default CSV;
export {CSV};
