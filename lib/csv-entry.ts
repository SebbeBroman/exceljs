/**
 * INTERNAL — not a package export.
 *
 * Side-effect entry that attaches CSV to the legacy Doc Workbook class for
 * vitest / historical tests (verquire). Public CSV API is named `csv` from
 * `@sebbebroman/exceljs` (see lib/csv/public.ts).
 *
 *   import {enableCsv} from '../lib/csv-entry.js'; // tests only
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
