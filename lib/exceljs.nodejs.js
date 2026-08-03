/**
 * Legacy namespace assembly (kept for deep imports).
 * Prefer package root `excel.js` named exports.
 */
import Workbook from './doc/workbook.js';
import ModelContainer from './doc/modelcontainer.js';
import WorkbookWriter from './stream/xlsx/workbook-writer.js';
import WorkbookReader from './stream/xlsx/workbook-reader.js';
import enums from './doc/enums.js';

const ExcelJS = {
  Workbook,
  ModelContainer,
  stream: {
    xlsx: {
      WorkbookWriter,
      WorkbookReader,
    },
  },
  ...enums,
};

export default ExcelJS;
export {ExcelJS, Workbook, ModelContainer, WorkbookWriter, WorkbookReader};
