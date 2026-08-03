// Browser entry without legacy core-js polyfill side-effects.
// Apps should provide their own polyfills if targeting old browsers.
export {default, ExcelJS, Workbook} from './exceljs.bare.js';
