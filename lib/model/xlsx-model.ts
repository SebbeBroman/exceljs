import type {ValueTypeCode} from './enums.js';
import type {
  Style,
  WorksheetProtection,
  WorksheetProperties,
  WorksheetState,
  PageSetup,
  HeaderFooter,
  RowBreak,
  WorksheetView,
  AutoFilter,
  ConditionalFormattingOptions,
  DataValidation,
} from './schema.js';
import type {TableData} from './table.js';
import type Range from './range.js';
type DataValidationsModel = Record<string, DataValidation>;
export interface RowModelCell {
  type: ValueTypeCode | number;
  address?: string;
  [key: string]: unknown;
}

export interface RowModelData {
  cells: RowModelCell[];
  number: number;
  min: number;
  max: number;
  height?: number;
  style?: Partial<Style>;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface ColumnModel {
  min: number;
  max: number;
  width: number;
  style?: Partial<Style>;
  isCustomWidth?: boolean;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface WorksheetSheetProtection extends Partial<WorksheetProtection> {
  sheet?: boolean;
  algorithmName?: string;
  saltValue?: string;
  hashValue?: string;
  spinCount?: number;
}

export interface WorksheetModelData {
  id?: number;
  name?: string;
  dataValidations?: DataValidationsModel;
  properties?: Partial<WorksheetProperties>;
  state?: WorksheetState;
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  rowBreaks?: RowBreak[];
  views?: WorksheetView[];
  autoFilter?: AutoFilter | null;
  media?: unknown[];
  sheetProtection?: WorksheetSheetProtection | null;
  tables?: TableData[];
  pivotTables?: unknown[];
  conditionalFormattings?: ConditionalFormattingOptions[];
  cols?: ColumnModel[];
  rows?: RowModelData[];
  dimensions?: Range;
  merges?: string[];
  mergeCells?: string[];
}
