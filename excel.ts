/**
 * Package entry for @sebbebroman/exceljs (builder-first API).
 *
 *   import { workbook, writeBuffer, load, csv } from '@sebbebroman/exceljs';
 *
 *   const buf = await workbook()
 *     .sheet('Data')
 *     .rows([['a', 1], ['b', 2]])
 *     .writeBuffer();
 *
 *   const data = await load(buf);
 *   const out = await workbook(data).sheet('Data').cell('A1', 'updated').writeBuffer();
 *
 *   const text = await workbook().sheet('Data', await csv.parse('a,1\nb,2')).csv();
 *
 * Node filesystem helpers: `@sebbebroman/exceljs/node`
 */

export {workbook, isWorkbookBuilder} from './lib/builder/workbook-builder.js';
export type {WorkbookBuilder, SheetBuilder} from './lib/builder/workbook-builder.js';

export {writeBuffer} from './lib/xlsx/write-buffer.js';
export {load} from './lib/xlsx/load.js';

/** Read-only view: `viewWorkbook` → `.sheet()` → `.rows()` / `.records()`. */
export {viewWorkbook, isWorkbookView} from './lib/read/view.js';
export type {
  WorkbookView,
  SheetView,
  ViewWorkbookOptions,
  RowsOptions,
  RecordsOptions,
} from './lib/read/view.js';

/** CSV/xlsx → dense `string[][]` (sugar over viewWorkbook). */
export {readRows, cellToDisplayString} from './lib/read/read-rows.js';
export type {ReadRowsOptions, ReadRowsFormat} from './lib/read/read-rows.js';

/** Named CSV helpers (`csv.parse` / `csv.stringify`). Tree-shaken when unused. */
export {csv, parseCsv, stringifyCsv} from './lib/csv/public.js';
export type {CsvParseOptions, CsvStringifyOptions} from './lib/csv/public.js';

export type {
  Workbook,
  WorkbookInit,
  WorkbookMeta,
  SheetModel,
  SheetRow,
  SheetCell,
  SheetInit,
  CellValue,
  Style,
  ColumnInput,
  RowInput,
  WriteOptions,
  LoadOptions,
  PageSetup,
  HeaderFooter,
  DataValidation,
  ConditionalFormattingOptions,
  TableProperties,
  NoteValue,
  MediaImage,
  SheetImageRange,
  ProtectOptions,
  SheetProtection,
  DefinedNameEntry,
  WorksheetViewInput,
  SheetTitleInput,
} from './lib/model/types.js';

// Enums still useful for formula / value discrimination when reading later
import enums from './lib/model/enums.js';

export const {ValueType, FormulaType, RelationshipType, DocumentType, ReadingOrder, ErrorValue} =
  enums;
export {enums};
