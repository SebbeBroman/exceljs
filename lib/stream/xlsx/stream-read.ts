/**
 * Public streaming read API for Node (`streamRead` from `@sebbebroman/exceljs/node`).
 * Thin wrapper around the internal WorkbookReader (not re-exported).
 */

import type {Readable} from 'node:stream';
import WorkbookReader, {
  type WorkbookReaderInput,
  type WorkbookStreamReaderOptions,
} from './workbook-reader.js';

export type StreamReadOptions = Partial<WorkbookStreamReaderOptions>;

/** One streamed data row from `streamRead`. */
export interface StreamReadRow {
  sheetName: string;
  /** 1-based sheet id from the reader. */
  sheetId: number | string;
  /** 1-based Excel row number. */
  rowNumber: number;
  /**
   * Cell values. Sparse array style from the legacy reader: index 0 is unused;
   * cells start at index 1 (Excel column A = 1).
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  values: any[];
}

/**
 * Async-iterate rows from an xlsx file path or readable stream without loading
 * the whole workbook into a plain model.
 *
 * @example
 * ```ts
 * for await (const { sheetName, rowNumber, values } of streamRead('big.xlsx')) {
 *   // values[1] is column A
 * }
 * ```
 */
export async function* streamRead(
  input: string | Readable | WorkbookReaderInput,
  options?: StreamReadOptions,
): AsyncGenerator<StreamReadRow> {
  const reader = new WorkbookReader(input, options);
  for await (const worksheetReader of reader) {
    for await (const row of worksheetReader) {
      yield {
        sheetName: worksheetReader.name,
        sheetId: worksheetReader.id,
        rowNumber: row.number,
        values: row.values,
      };
    }
  }
}
