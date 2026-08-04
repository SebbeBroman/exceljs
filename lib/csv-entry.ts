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

// Workbook may already declare a CSV static; attach our implementation at runtime.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function enableCsv(WorkbookClass: any = Workbook): typeof Workbook {
  WorkbookClass.CSV = CSV;
  return WorkbookClass;
}

// Auto-enable when this entry is imported
enableCsv(Workbook);

export {CSV};
export default CSV;
