/**
 * INTERNAL — not a package export.
 *
 * Browser-oriented bare assembly of the legacy Doc Workbook class (no Node
 * stream/fs). Kept for internal experiments; public API is builder-first via
 * excel.ts / node.ts.
 */
import Workbook from './doc/workbook.js';
import enums from './doc/enums.js';

const ExcelJS = {
  Workbook,
  ...enums,
};

export default ExcelJS;
export {ExcelJS, Workbook};
