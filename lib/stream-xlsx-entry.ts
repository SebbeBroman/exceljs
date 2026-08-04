/**
 * Streaming XLSX API entry — import only when needed so the main package can tree-shake it.
 *
 *   import {WorkbookWriter, WorkbookReader} from 'exceljs/stream/xlsx';
 */
export {default as WorkbookWriter} from './stream/xlsx/workbook-writer.js';
export {default as WorkbookReader} from './stream/xlsx/workbook-reader.js';
export {default as WorksheetWriter} from './stream/xlsx/worksheet-writer.js';
export {default as WorksheetReader} from './stream/xlsx/worksheet-reader.js';
