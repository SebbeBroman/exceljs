/**
 * Optional CSV support entry.
 * Side-effect: enables workbook.csv on Workbook.
 *
 *   import 'exceljs/csv';
 *   // or
 *   import {enableCsv, CSV} from 'exceljs/csv';
 */
import CSV from './csv/csv.js';
import Workbook from './doc/workbook.js';

export function enableCsv(WorkbookClass = Workbook) {
  WorkbookClass.CSV = CSV;
  return WorkbookClass;
}

// Auto-enable when this entry is imported
enableCsv(Workbook);

export {CSV};
export default CSV;
