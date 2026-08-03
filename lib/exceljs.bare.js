// Browser-oriented bare entry (no polyfills) — ESM
import Workbook from './doc/workbook.js';
import enums from './doc/enums.js';

const ExcelJS = {
  Workbook,
  ...enums,
};

export default ExcelJS;
export {ExcelJS, Workbook};
