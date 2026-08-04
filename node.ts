/**
 * Node-only adapters for @sebbebroman/excel-ts.
 *
 *   import {
 *     workbook, writeFile, readFile, streamWrite, streamRead,
 *     readCsvFile, writeCsvFile,
 *   } from '@sebbebroman/excel-ts/node';
 *
 *   await writeFile('out.xlsx', workbook().sheet('A').row([1, 2]));
 *   await streamWrite('big.xlsx', { sheets: [{ name: 'Data', rows: source }] });
 *   const data = await readFile('out.xlsx');
 *   const fromCsv = await readCsvFile('data.csv');
 *   await writeCsvFile('out.csv', fromCsv);
 */

import fs from 'node:fs/promises';
import type {Workbook, WriteOptions} from './lib/model/types.js';
import type {WorkbookBuilder} from './lib/builder/workbook-builder.js';
import {writeBuffer} from './lib/xlsx/write-buffer.js';
import {load} from './lib/xlsx/load.js';
import type {LoadOptions} from './lib/model/types.js';
import {csv, type CsvParseOptions, type CsvStringifyOptions} from './lib/csv/public.js';
import {workbook as createWorkbook} from './lib/builder/workbook-builder.js';

export {
  workbook,
  isWorkbookBuilder,
  writeBuffer,
  load,
  viewWorkbook,
  isWorkbookView,
  readRows,
  cellToDisplayString,
  csv,
  parseCsv,
  stringifyCsv,
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
  enums,
} from './excel.js';

export type {
  Workbook,
  WorkbookInit,
  WorkbookMeta,
  SheetModel,
  SheetInit,
  CellValue,
  Style,
  ColumnInput,
  RowInput,
  WriteOptions,
  LoadOptions,
  ReadRowsOptions,
  ReadRowsFormat,
  WorkbookView,
  SheetView,
  ViewWorkbookOptions,
  RowsOptions,
  RecordsOptions,
  WorkbookBuilder,
  SheetBuilder,
  CsvParseOptions,
  CsvStringifyOptions,
} from './excel.js';

export {streamWrite} from './lib/stream/xlsx/stream-write.js';
export type {
  StreamWriteOptions,
  StreamSheetOptions,
  StreamWriteSheetSpec,
  StreamWriteDeclarative,
  StreamWriteSpec,
  StreamWriteCallback,
  StreamSheetHandle,
  StreamWorkbookHandle,
} from './lib/stream/xlsx/stream-write.js';

export {streamRead} from './lib/stream/xlsx/stream-read.js';
export type {StreamReadOptions, StreamReadRow} from './lib/stream/xlsx/stream-read.js';
/**
 * Write a builder or plain workbook to a filesystem path.
 */
export async function writeFile(
  path: string,
  input: Workbook | WorkbookBuilder,
  options?: WriteOptions,
): Promise<void> {
  const bytes = await writeBuffer(input, options);
  await fs.writeFile(path, bytes);
}

/**
 * Read an xlsx file from disk into a plain Workbook snapshot.
 */
export async function readFile(path: string, opts?: LoadOptions): Promise<Workbook> {
  const bytes = await fs.readFile(path);
  return load(bytes, opts);
}

/**
 * Read a CSV file into a plain Workbook (single sheet).
 */
export async function readCsvFile(path: string, opts?: CsvParseOptions): Promise<Workbook> {
  const text = await fs.readFile(path, 'utf8');
  const init = await csv.parse(text, opts);
  const sheetName = opts?.sheetName ?? 'Sheet1';
  return createWorkbook().sheet(sheetName, init).build();
}

/**
 * Write a builder or plain workbook sheet as CSV to a filesystem path.
 */
export async function writeCsvFile(
  path: string,
  input: Workbook | WorkbookBuilder,
  opts?: CsvStringifyOptions,
): Promise<void> {
  const text = await csv.stringify(input, opts);
  await fs.writeFile(path, text, opts?.encoding ?? 'utf8');
}
