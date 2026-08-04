/**
 * Package entry for @sebbebroman/excel-ts.
 *
 * Prefer named imports for tree-shaking:
 *   import { Workbook } from '@sebbebroman/excel-ts';
 *
 * Optional features:
 *   import '@sebbebroman/excel-ts/csv';
 *   import { WorkbookWriter } from '@sebbebroman/excel-ts/stream/xlsx';
 */

import Workbook from './lib/doc/workbook.js';
import ModelContainer from './lib/doc/modelcontainer.js';
import enums from './lib/doc/enums.js';
import {ensureDocFeatures} from './lib/doc/doc-features.js';

export {Workbook, ModelContainer, ensureDocFeatures};

// Stream APIs live at '@sebbebroman/excel-ts/stream/xlsx' so the main entry stays lean.

// Enum values as named exports (matches index.d.ts usage patterns)
export const {ValueType, FormulaType, RelationshipType, DocumentType, ReadingOrder, ErrorValue} =
  enums;

export {enums};

// Namespace-style default (ExcelJS-compatible shape)
// Note: does not include stream.* or csv — use the dedicated entry points
const ExcelJS = {
  Workbook,
  ModelContainer,
  ensureDocFeatures,
  ValueType: enums.ValueType,
  FormulaType: enums.FormulaType,
  RelationshipType: enums.RelationshipType,
  DocumentType: enums.DocumentType,
  ReadingOrder: enums.ReadingOrder,
  ErrorValue: enums.ErrorValue,
};

export default ExcelJS;
