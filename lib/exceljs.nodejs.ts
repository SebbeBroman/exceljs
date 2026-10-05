/**
 * INTERNAL — not a package export.
 *
 * Legacy namespace assembly for vitest (verquire) and historical tests that
 * still exercise the Doc Workbook / stream writer classes. Public package
 * surface is builder-first: excel.ts + node.ts only.
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
