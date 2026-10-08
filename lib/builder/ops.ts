import type {
  CellValue,
  ColumnInput,
  ConditionalFormattingOptions,
  DataValidation,
  HeaderFooter,
  MediaImage,
  NoteValue,
  PageSetup,
  RowInput,
  SheetImageRange,
  Style,
  TableProperties,
  WorkbookMeta,
  WorksheetViewInput,
} from '../model/types.js';

export interface UsedFlags {
  styles: boolean;
  merges: boolean;
  formulas: boolean;
  columns: boolean;
  views: boolean;
  pageSetup: boolean;
  dataValidations: boolean;
  conditionalFormatting: boolean;
  notes: boolean;
  protection: boolean;
  tables: boolean;
  images: boolean;
  definedNames: boolean;
}

export function emptyUsedFlags(): UsedFlags {
  return {
    styles: false,
    merges: false,
    formulas: false,
    columns: false,
    views: false,
    pageSetup: false,
    dataValidations: false,
    conditionalFormatting: false,
    notes: false,
    protection: false,
    tables: false,
    images: false,
    definedNames: false,
  };
}

export type BuilderOp =
  | {op: 'meta'; meta: WorkbookMeta}
  | {op: 'sheet'; name: string}
  | {op: 'columns'; sheet: string; columns: ColumnInput[]}
  | {op: 'row'; sheet: string; values: RowInput}
  | {op: 'rows'; sheet: string; values: RowInput[]}
  | {op: 'cell'; sheet: string; address: string; value: CellValue; style?: Style}
  | {op: 'cells'; sheet: string; map: Record<string, CellValue>}
  | {op: 'style'; sheet: string; range: string; style: Style}
  | {op: 'merge'; sheet: string; range: string}
  // Phase 5 — advanced sheet / workbook features
  | {op: 'views'; sheet: string; views: WorksheetViewInput[]}
  | {op: 'pageSetup'; sheet: string; pageSetup: Partial<PageSetup>}
  | {op: 'headerFooter'; sheet: string; headerFooter: Partial<HeaderFooter>}
  | {op: 'dataValidation'; sheet: string; address: string; rules: DataValidation}
  | {op: 'conditionalFormatting'; sheet: string; cf: ConditionalFormattingOptions}
  | {op: 'note'; sheet: string; address: string; note: NoteValue}
  /** Loaded / pre-hashed protection model applied as-is at materialize. */
  | {op: 'sheetProtection'; sheet: string; model: Record<string, unknown>}
  | {op: 'table'; sheet: string; table: TableProperties}
  | {op: 'media'; id: number; image: MediaImage}
  | {op: 'sheetImage'; sheet: string; imageId: number; range: SheetImageRange}
  | {op: 'definedName'; name: string; refersTo: string};

export function noteFormulaValue(value: CellValue, flags: UsedFlags): void {
  if (value && typeof value === 'object' && 'formula' in value) {
    flags.formulas = true;
  }
  if (value && typeof value === 'object' && 'sharedFormula' in value) {
    flags.formulas = true;
  }
}
