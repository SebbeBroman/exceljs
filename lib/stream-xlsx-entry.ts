/**
 * INTERNAL — not a package export.
 *
 * Re-exports legacy stream WorkbookWriter / WorkbookReader for deep imports
 * used by internal benches/tests. Public streaming API is `streamWrite` /
 * `streamRead` on `@sebbebroman/exceljs/node`.
 */
export {default as WorkbookWriter} from './stream/xlsx/workbook-writer.js';
export {default as WorkbookReader} from './stream/xlsx/workbook-reader.js';
export {default as WorksheetWriter} from './stream/xlsx/worksheet-writer.js';
export {default as WorksheetReader} from './stream/xlsx/worksheet-reader.js';
