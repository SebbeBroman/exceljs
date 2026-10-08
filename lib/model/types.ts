/**
 * Plain (immutable-friendly) workbook model for the builder API.
 * This is the public data shape returned by `build()`; it is not the mutable ExcelJS class graph.
 */

import type {
  Alignment,
  Borders,
  CellErrorValue,
  CellFormulaValue,
  CellHyperlinkValue,
  CellRichTextValue,
  CellSharedFormulaValue,
  Comment,
  ConditionalFormattingOptions as IndexCFOptions,
  DataValidation as IndexDataValidation,
  Fill,
  Font,
  HeaderFooter as IndexHeaderFooter,
  PageSetup as IndexPageSetup,
  Protection,
  Style as FullStyle,
  TableProperties as IndexTableProperties,
  WorkbookProperties,
  WorkbookView,
  WorksheetProtection,
  WorksheetView,
} from './schema.js';

/** Partial style patches accepted by the builder. */
export type Style = Partial<
  Pick<FullStyle, 'numFmt'> & {
    font: Partial<Font>;
    alignment: Partial<Alignment>;
    protection: Partial<Protection>;
    border: Partial<Borders>;
    fill: Fill;
  }
>;

export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | undefined
  | CellErrorValue
  | CellRichTextValue
  | CellHyperlinkValue
  | CellFormulaValue
  | CellSharedFormulaValue;

export interface WorkbookMeta {
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
  company?: string;
  manager?: string;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  language?: string;
  revision?: Date | string | number;
  contentStatus?: string;
  properties?: Partial<WorkbookProperties>;
  views?: WorkbookView[];
}

export interface ColumnInput {
  /**
   * Header label. Only the first line is used when `header` is an array
   * (multi-row headers are not supported — extra lines are dropped
   * consistently across build/writeBuffer/stream paths).
   */
  header?: string | string[];
  key?: string;
  width?: number;
  hidden?: boolean;
  style?: Style;
  outlineLevel?: number;
}

/** Array of cell values, or a keyed object when columns define keys. */
export type RowInput = ReadonlyArray<CellValue> | Record<string, CellValue>;

/** Title row above a table: plain text, or text with optional style/merge.
 * `style` applies to the `merge` range when given, otherwise to the title
 * cell (A1 of the emitted row). `merge` must be a valid A1 range.
 */
export type SheetTitleInput = string | {text: string; style?: Style; merge?: string};

export interface SheetCell {
  value: CellValue;
  style?: Style;
}

export interface SheetRow {
  number: number;
  cells: Record<number, SheetCell>; // 1-based col → cell
  height?: number;
  hidden?: boolean;
  style?: Style;
}

// --- Phase 5 feature types (re-exported / aliased from legacy index for builder) ---

export type PageSetup = IndexPageSetup;
export type HeaderFooter = IndexHeaderFooter;
export type DataValidation = IndexDataValidation;
export type ConditionalFormattingOptions = IndexCFOptions;
export type TableProperties = IndexTableProperties;
export type NoteValue = string | Comment;
export type WorksheetViewInput = Partial<WorksheetView>;
export type ProtectOptions = Partial<WorksheetProtection> & {spinCount?: number};

/** Workbook media image for `.image({ extension, … })`. */
export interface MediaImage {
  extension: 'jpeg' | 'png' | 'gif';
  base64?: string;
  filename?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  buffer?: any;
}

/** Range / position for placing a sheet image. */
export type SheetImageRange =
  | string
  | {
      tl: {col: number; row: number};
      br: {col: number; row: number};
      editAs?: string;
      hyperlinks?: {hyperlink: string; tooltip?: string};
    }
  | {
      tl: {col: number; row: number};
      ext: {width: number; height: number};
      editAs?: string;
      hyperlinks?: {hyperlink: string; tooltip?: string};
    };

/** Prepared OOXML sheet protection; create password hashes through the protection entry. */
export type SheetProtection = ProtectOptions & {
  sheet?: boolean;
  algorithmName?: string;
  saltValue?: string;
  hashValue?: string;
  [key: string]: unknown;
};

/** Defined name entry for plain model / builder. */
export interface DefinedNameEntry {
  name: string;
  /** Range or formula, e.g. `Sheet1!$A$1` or `Sheet1!$A$1:$B$2`. */
  refersTo: string;
}

export interface SheetImagePlacement {
  imageId: number;
  range: SheetImageRange;
}

export interface SheetModel {
  id: number;
  name: string;
  rows: SheetRow[];
  columns?: ColumnInput[];
  merges?: string[];
  /** Sheet view(s), e.g. freeze panes. */
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  /** Address (or range) → validation rules. */
  dataValidations?: Record<string, DataValidation>;
  conditionalFormattings?: ConditionalFormattingOptions[];
  /** Cell address → note text or comment object. */
  notes?: Record<string, NoteValue>;
  /**
   * Pre-hashed OOXML protection model (e.g. from load). Applied as-is on write.
   * Write-only for re-encode; password cannot be recovered.
   */
  sheetProtection?: SheetProtection;
  tables?: TableProperties[];
  /** Sheet-level image placements (media is on Workbook). */
  images?: SheetImagePlacement[];
}

export interface Workbook {
  meta: WorkbookMeta;
  sheets: SheetModel[];
  /** Workbook media (images). Indices are image ids used by sheet placements. */
  media?: MediaImage[];
  definedNames?: DefinedNameEntry[];
}

export interface WorkbookInit extends WorkbookMeta {
  sheets?: SheetModel[];
  media?: MediaImage[];
  definedNames?: DefinedNameEntry[];
}

export interface WriteOptions {
  useSharedStrings?: boolean;
  useStyles?: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  zip?: any;
}

/** Options for `load` / Node `readFile`. */
export interface LoadOptions {
  /** Treat input as base64 when data is a string. */
  base64?: boolean;
  /** XML node names to ignore while parsing. */
  ignoreNodes?: string[];
}

export interface SheetInit {
  /** Optional title row emitted before columns/rows. */
  title?: SheetTitleInput;
  columns?: ColumnInput[];
  rows?: RowInput[];
  merges?: string[];
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
}
