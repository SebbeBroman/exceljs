/**
 * ExcelJS ESM entry (Vite / SvelteKit / modern bundlers).
 *
 * Prefer named imports for tree-shaking:
 *   import { Workbook } from 'exceljs';
 *
 * Optional features:
 *   import 'exceljs/csv';
 *   import { WorkbookWriter } from 'exceljs/stream/xlsx';
 */

import Workbook from './lib/doc/workbook.js';
import ModelContainer from './lib/doc/modelcontainer.js';
import enums from './lib/doc/enums.js';

export {Workbook, ModelContainer};

// Stream APIs live at 'exceljs/stream/xlsx' so the main entry stays lean.

// Enum values as named exports (matches index.d.ts usage patterns)
export const {
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
} = enums;

export {enums};

// Namespace-style default for drop-in migration from `const ExcelJS = require('exceljs')`
// Note: does not include stream.* (import from 'exceljs/stream/xlsx') or csv (import 'exceljs/csv')
const ExcelJS = {
  Workbook,
  ModelContainer,
  ValueType: enums.ValueType,
  FormulaType: enums.FormulaType,
  RelationshipType: enums.RelationshipType,
  DocumentType: enums.DocumentType,
  ReadingOrder: enums.ReadingOrder,
  ErrorValue: enums.ErrorValue,
};

export default ExcelJS;
