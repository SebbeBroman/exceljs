/**
 * Public types for @sebbebroman/exceljs 0.1.x (builder-first).
 *
 * Package `"types"` and `exports["."].types` resolve here.
 * Self-contained declarations for the supported public API.
 */

// Shared shapes used by the public workbook model and builder.
declare enum LegacyRelationshipType {
  None = 0,
  OfficeDocument = 1,
  Worksheet = 2,
  CalcChain = 3,
  SharedStrings = 4,
  Styles = 5,
  Theme = 6,
  Hyperlink = 7,
}

declare enum LegacyDocumentType {
  Xlsx = 1,
}

declare const enum PaperSize {
  Legal = 5,
  Executive = 7,
  A4 = 9,
  A5 = 11,
  B5 = 13,
  Envelope_10 = 20,
  Envelope_DL = 27,
  Envelope_C5 = 28,
  Envelope_B5 = 34,
  Envelope_Monarch = 37,
  Double_Japan_Postcard_Rotated = 82,
  K16_197x273_mm = 119,
}

interface WorksheetViewCommon {
  /**
   * Sets the worksheet view's orientation to right-to-left, `false` by default
   */
  rightToLeft: boolean;

  /**
   * The currently selected cell
   */
  activeCell: string;

  /**
   * Shows or hides the ruler in Page Layout, `true` by default
   */
  showRuler: boolean;

  /**
   * Shows or hides the row and column headers (e.g. A1, B1 at the top and 1,2,3 on the left,
   * `true` by default
   */
  showRowColHeaders: boolean;

  /**
   * Shows or hides the gridlines (shown for cells where borders have not been defined),
   * `true` by default
   */
  showGridLines: boolean;

  /**
   * 	Percentage zoom to use for the view, `100` by default
   */
  zoomScale: number;

  /**
   * 	Normal zoom for the view, `100` by default
   */
  zoomScaleNormal: number;
}

interface WorksheetViewNormal {
  /**
   * Controls the view state
   */
  state: 'normal';

  /**
   * Presentation style
   */
  style: 'pageBreakPreview' | 'pageLayout';
}

interface WorksheetViewFrozen {
  /**
   * Where a number of rows and columns to the top and left are frozen in place.
   * Only the bottom left section will scroll
   */
  state: 'frozen';

  /**
   * Presentation style
   */
  style?: 'pageBreakPreview';

  /**
   * How many columns to freeze. To freeze rows only, set this to 0 or undefined
   */
  xSplit?: number;

  /**
   * How many rows to freeze. To freeze columns only, set this to 0 or undefined
   */
  ySplit?: number;

  /**
   * Which cell will be top-left in the bottom-right pane. Note: cannot be a frozen cell.
   * Defaults to first unfrozen cell
   */
  topLeftCell?: string;
}

interface WorksheetViewSplit {
  /**
   * Where the view is split into 4 sections, each semi-independently scrollable.
   */
  state: 'split';

  /**
   * Presentation style
   */
  style?: 'pageBreakPreview' | 'pageLayout';

  /**
   * How many points from the left to place the splitter.
   * To split vertically, set this to 0 or undefined
   */
  xSplit?: number;

  /**
   * How many points from the top to place the splitter.
   * To split horizontally, set this to 0 or undefined
   */
  ySplit?: number;

  /**
   * Which cell will be top-left in the bottom-right pane
   */
  topLeftCell?: string;

  /**
   * Which pane will be active
   */
  activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
}

type WorksheetView = WorksheetViewCommon &
  (WorksheetViewNormal | WorksheetViewFrozen | WorksheetViewSplit);

interface LegacyWorkbookView {
  x: number;
  y: number;
  width: number;
  height: number;
  firstSheet: number;
  activeTab: number;
  visibility: string;
}

type FillPatterns =
  | 'none'
  | 'solid'
  | 'darkVertical'
  | 'darkHorizontal'
  | 'darkGrid'
  | 'darkTrellis'
  | 'darkDown'
  | 'darkUp'
  | 'lightVertical'
  | 'lightHorizontal'
  | 'lightGrid'
  | 'lightTrellis'
  | 'lightDown'
  | 'lightUp'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray125'
  | 'gray0625';

interface FillPattern {
  type: 'pattern';
  pattern: FillPatterns;
  fgColor?: Partial<Color>;
  bgColor?: Partial<Color>;
}

interface GradientStop {
  position: number;
  color: Partial<Color>;
}

interface FillGradientAngle {
  type: 'gradient';
  gradient: 'angle';

  /**
   * For 'angle' gradient, specifies the direction of the gradient. 0 is from the left to the right.
   * Values from 1 - 359 rotates the direction clockwise
   */
  degree: number;

  /**
   * Specifies the gradient colour sequence. Is an array of objects containing position and
   * color starting with position 0 and ending with position 1.
   * Intermediary positions may be used to specify other colours on the path.
   */
  stops: GradientStop[];
}

interface FillGradientPath {
  type: 'gradient';
  gradient: 'path';

  /**
   * For 'path' gradient. Specifies the relative coordinates for the start of the path.
   * 'left' and 'top' values range from 0 to 1
   */
  center: {left: number; top: number};

  /**
   * Specifies the gradient colour sequence. Is an array of objects containing position and
   * color starting with position 0 and ending with position 1.
   * Intermediary positions may be used to specify other colours on the path.
   */
  stops: GradientStop[];
}

type Fill = FillPattern | FillGradientAngle | FillGradientPath;

interface Font {
  name: string;
  size: number;
  family: number;
  scheme: 'minor' | 'major' | 'none';
  charset: number;
  color: Partial<Color>;
  bold: boolean;
  italic: boolean;
  underline: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  vertAlign: 'superscript' | 'subscript';
  strike: boolean;
  outline: boolean;
}

type BorderStyle =
  | 'thin'
  | 'dotted'
  | 'hair'
  | 'medium'
  | 'double'
  | 'thick'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'slantDashDot'
  | 'mediumDashed'
  | 'mediumDashDotDot'
  | 'mediumDashDot';

interface Color {
  /**
   * Hex string for alpha-red-green-blue e.g. FF00FF00
   */
  argb: string;

  /**
   * Choose a theme by index
   */
  theme: number;
}

interface Border {
  style: BorderStyle;
  color: Partial<Color>;
}

interface BorderDiagonal extends Border {
  up: boolean;
  down: boolean;
}

interface Borders {
  top: Partial<Border>;
  left: Partial<Border>;
  bottom: Partial<Border>;
  right: Partial<Border>;
  diagonal: Partial<BorderDiagonal>;
}

interface Margins {
  top: number;
  left: number;
  bottom: number;
  right: number;
  header: number;
  footer: number;
}

declare enum LegacyReadingOrder {
  LeftToRight = 1,
  RightToLeft = 2,
}

interface Alignment {
  horizontal: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
  wrapText: boolean;
  shrinkToFit: boolean;
  indent: number;
  readingOrder: 'rtl' | 'ltr';
  textRotation: number | 'vertical';
}

interface Protection {
  locked: boolean;
  hidden: boolean;
}

interface FullStyle {
  numFmt: string;
  font: Partial<Font>;
  alignment: Partial<Alignment>;
  protection: Partial<Protection>;
  border: Partial<Borders>;
  fill: Fill;
}

type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

interface DataValidation {
  type: 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom';
  formulae: any[];
  allowBlank?: boolean;
  operator?: DataValidationOperator;
  error?: string;
  errorTitle?: string;
  errorStyle?: string;
  prompt?: string;
  promptTitle?: string;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
}

declare enum LegacyErrorValue {
  NotApplicable = '#N/A',
  Ref = '#REF!',
  Name = '#NAME?',
  DivZero = '#DIV/0!',
  Null = '#NULL!',
  Value = '#VALUE!',
  Num = '#NUM!',
}

interface CellErrorValue {
  error: '#N/A' | '#REF!' | '#NAME?' | '#DIV/0!' | '#NULL!' | '#VALUE!' | '#NUM!';
}

interface RichText {
  text: string;
  font?: Partial<Font>;
}

interface CellRichTextValue {
  richText: RichText[];
}

interface CellHyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

interface CellFormulaValue {
  formula: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

interface CellSharedFormulaValue {
  sharedFormula: string;
  readonly formula?: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

declare enum LegacyValueType {
  Null = 0,
  Merge = 1,
  Number = 2,
  String = 3,
  Date = 4,
  Hyperlink = 5,
  Formula = 6,
  SharedString = 7,
  RichText = 8,
  Boolean = 9,
  Error = 10,
}

declare enum LegacyFormulaType {
  None = 0,
  Master = 1,
  Shared = 2,
}

interface CommentMargins {
  insetmode: 'auto' | 'custom';
  inset: number[];
}

interface CommentProtection {
  locked: 'True' | 'False';
  lockText: 'True' | 'False';
}

type CommentEditAs = 'twoCells' | 'oneCells' | 'absolute';

interface Comment {
  texts?: RichText[];
  margins?: Partial<CommentMargins>;
  protection?: Partial<CommentProtection>;
  editAs?: CommentEditAs;
}

interface PageSetup {
  /**
   * Whitespace on the borders of the page. Units are inches.
   */
  margins: Margins;

  /**
   * Orientation of the page - i.e. taller (`'portrait'`) or wider (`'landscape'`).
   *
   * `'portrait'` by default
   */
  orientation: 'portrait' | 'landscape';

  /**
   * Horizontal Dots per Inch. Default value is 4294967295
   */
  horizontalDpi: number;

  /**
   * Vertical Dots per Inch. Default value is 4294967295
   */
  verticalDpi: number;

  /**
   * Whether to use fitToWidth and fitToHeight or scale settings.
   *
   * Default is based on presence of these settings in the pageSetup object - if both are present,
   * scale wins (i.e. default will be false)
   */
  fitToPage: boolean;

  /**
   * How many pages wide the sheet should print on to. Active when fitToPage is true
   *
   * Default is 1
   */
  fitToWidth: number;

  /**
   * How many pages high the sheet should print on to. Active when fitToPage is true
   *
   * Default is 1
   */
  fitToHeight: number;

  /**
   * Percentage value to increase or reduce the size of the print. Active when fitToPage is false
   *
   * Default is 100
   */
  scale: number;

  /**
   * Which order to print the pages.
   *
   * Default is `downThenOver`
   */
  pageOrder: 'downThenOver' | 'overThenDown';

  /**
   * Print without colour
   *
   * false by default
   */
  blackAndWhite: boolean;

  /**
   * Print with less quality (and ink)
   *
   * false by default
   */
  draft: boolean;

  /**
   * Where to place comments
   *
   * Default is `None`
   */
  cellComments: 'atEnd' | 'asDisplayed' | 'None';

  /**
   * Where to show errors
   *
   * Default is `displayed`
   */
  errors: 'dash' | 'blank' | 'NA' | 'displayed';

  /**
   * 	What paper size to use (see below)
   *
   * | Name                          | Value       |
   * | ----------------------------- | ---------   |
   * | Letter                        | `undefined` |
   * | Legal                         |  `5`        |
   * | Executive                     |  `7`        |
   * | A4                            |  `9`        |
   * | A5                            |  `11`       |
   * | B5 (JIS)                      |  `13`       |
   * | Envelope #10                  |  `20`       |
   * | Envelope DL                   |  `27`       |
   * | Envelope C5                   |  `28`       |
   * | Envelope B5                   |  `34`       |
   * | Envelope Monarch              |  `37`       |
   * | Double Japan Postcard Rotated |  `82`       |
   * | 16K 197x273 mm                |  `119`      |
   */
  paperSize: PaperSize;

  /**
   * Whether to show the row numbers and column letters, `false` by default
   */
  showRowColHeaders: boolean;

  /**
   * Whether to show grid lines, `false` by default
   */
  showGridLines: boolean;

  /**
   * Which number to use for the first page
   */
  firstPageNumber: number;

  /**
   * 	Whether to center the sheet data horizontally, `false` by default
   */
  horizontalCentered: boolean;

  /**
   * 	Whether to center the sheet data vertically, `false` by default
   */
  verticalCentered: boolean;

  /**
   * Set Print Area for a sheet, e.g. `'A1:G20'`
   */
  printArea: string;

  /**
   * Repeat specific rows on every printed page, e.g. `'1:3'`
   */
  printTitlesRow: string;

  /**
   * Repeat specific columns on every printed page, e.g. `'A:C'`
   */
  printTitlesColumn: string;
}

interface HeaderFooter {
  /**
   * Set the value of differentFirst as true, which indicates that headers/footers for first page are different from the other pages, `false` by default
   */
  differentFirst: boolean;
  /**
   * Set the value of differentOddEven as true, which indicates that headers/footers for odd and even pages are different, `false` by default
   */
  differentOddEven: boolean;
  /**
   * Set header string for odd pages, could format the string and `null` by default
   */
  oddHeader: string;
  /**
   * Set footer string for odd pages, could format the string and `null` by default
   */
  oddFooter: string;
  /**
   * Set header string for even pages, could format the string and `null` by default
   */
  evenHeader: string;
  /**
   * Set footer string for even pages, could format the string and `null` by default
   */
  evenFooter: string;
  /**
   * Set header string for the first page, could format the string and `null` by default
   */
  firstHeader: string;
  /**
   * Set footer string for the first page, could format the string and `null` by default
   */
  firstFooter: string;
}

interface WorksheetProtection {
  objects: boolean;
  scenarios: boolean;
  selectLockedCells: boolean;
  selectUnlockedCells: boolean;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  insertHyperlinks: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
  pivotTables: boolean;
  spinCount: number;
}

type CellIsOperators = 'equal' | 'greaterThan' | 'lessThan' | 'between';

type ContainsTextOperators =
  | 'containsText'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors';

type TimePeriodTypes =
  | 'lastWeek'
  | 'thisWeek'
  | 'nextWeek'
  | 'yesterday'
  | 'today'
  | 'tomorrow'
  | 'last7Days'
  | 'lastMonth'
  | 'thisMonth'
  | 'nextMonth';

type IconSetTypes =
  | '5Arrows'
  | '5ArrowsGray'
  | '5Boxes'
  | '5Quarters'
  | '5Rating'
  | '4Arrows'
  | '4ArrowsGray'
  | '4Rating'
  | '4RedToBlack'
  | '4TrafficLights'
  | 'NoIcons'
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3Signs'
  | '3Stars'
  | '3Symbols'
  | '3Symbols2'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Triangles';

type CfvoTypes =
  | 'percentile'
  | 'percent'
  | 'num'
  | 'min'
  | 'max'
  | 'formula'
  | 'autoMin'
  | 'autoMax';

interface Cvfo {
  type: CfvoTypes;
  value?: number;
}

interface ConditionalFormattingBaseRule {
  priority: number;
  style?: Partial<FullStyle>;
}

interface ExpressionRuleType extends ConditionalFormattingBaseRule {
  type: 'expression';
  formulae?: any[];
}

interface CellIsRuleType extends ConditionalFormattingBaseRule {
  type: 'cellIs';
  formulae?: any[];
  operator?: CellIsOperators;
}

interface Top10RuleType extends ConditionalFormattingBaseRule {
  type: 'top10';
  rank: number;
  percent: boolean;
  bottom: boolean;
}

interface AboveAverageRuleType extends ConditionalFormattingBaseRule {
  type: 'aboveAverage';
  aboveAverage: boolean;
}

interface ColorScaleRuleType extends ConditionalFormattingBaseRule {
  type: 'colorScale';
  cfvo?: Cvfo[];
  color?: Partial<Color>[];
}

interface IconSetRuleType extends ConditionalFormattingBaseRule {
  type: 'iconSet';
  showValue?: boolean;
  reverse?: boolean;
  custom?: boolean;
  iconSet?: IconSetTypes;
  cfvo?: Cvfo[];
}

interface ContainsTextRuleType extends ConditionalFormattingBaseRule {
  type: 'containsText';
  operator?: ContainsTextOperators;
  text?: string;
}

interface TimePeriodRuleType extends ConditionalFormattingBaseRule {
  type: 'timePeriod';
  timePeriod?: TimePeriodTypes;
}

interface DataBarRuleType extends ConditionalFormattingBaseRule {
  type: 'dataBar';
  gradient?: boolean;
  minLength?: number;
  maxLength?: number;
  showValue?: boolean;
  border?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: 'auto' | 'middle' | 'none';
  direction?: 'context' | 'leftToRight' | 'rightToLeft';
  cfvo?: Cvfo[];
}

type ConditionalFormattingRule =
  | ExpressionRuleType
  | CellIsRuleType
  | Top10RuleType
  | AboveAverageRuleType
  | ColorScaleRuleType
  | IconSetRuleType
  | ContainsTextRuleType
  | TimePeriodRuleType
  | DataBarRuleType;

interface ConditionalFormattingOptions {
  ref: string;
  rules: ConditionalFormattingRule[];
}

interface WorkbookProperties {
  /**
   * Set workbook dates to 1904 date system
   */
  date1904: boolean;
}

interface TableStyleProperties {
  /**
   * The colour theme of the table
   * @default 'TableStyleMedium2'
   */
  theme?:
    | 'TableStyleDark1'
    | 'TableStyleDark10'
    | 'TableStyleDark11'
    | 'TableStyleDark2'
    | 'TableStyleDark3'
    | 'TableStyleDark4'
    | 'TableStyleDark5'
    | 'TableStyleDark6'
    | 'TableStyleDark7'
    | 'TableStyleDark8'
    | 'TableStyleDark9'
    | 'TableStyleLight1'
    | 'TableStyleLight10'
    | 'TableStyleLight11'
    | 'TableStyleLight12'
    | 'TableStyleLight13'
    | 'TableStyleLight14'
    | 'TableStyleLight15'
    | 'TableStyleLight16'
    | 'TableStyleLight17'
    | 'TableStyleLight18'
    | 'TableStyleLight19'
    | 'TableStyleLight2'
    | 'TableStyleLight20'
    | 'TableStyleLight21'
    | 'TableStyleLight3'
    | 'TableStyleLight4'
    | 'TableStyleLight5'
    | 'TableStyleLight6'
    | 'TableStyleLight7'
    | 'TableStyleLight8'
    | 'TableStyleLight9'
    | 'TableStyleMedium1'
    | 'TableStyleMedium10'
    | 'TableStyleMedium11'
    | 'TableStyleMedium12'
    | 'TableStyleMedium13'
    | 'TableStyleMedium14'
    | 'TableStyleMedium15'
    | 'TableStyleMedium16'
    | 'TableStyleMedium17'
    | 'TableStyleMedium18'
    | 'TableStyleMedium19'
    | 'TableStyleMedium2'
    | 'TableStyleMedium20'
    | 'TableStyleMedium21'
    | 'TableStyleMedium22'
    | 'TableStyleMedium23'
    | 'TableStyleMedium24'
    | 'TableStyleMedium25'
    | 'TableStyleMedium26'
    | 'TableStyleMedium27'
    | 'TableStyleMedium28'
    | 'TableStyleMedium3'
    | 'TableStyleMedium4'
    | 'TableStyleMedium5'
    | 'TableStyleMedium6'
    | 'TableStyleMedium7'
    | 'TableStyleMedium8'
    | 'TableStyleMedium9';
  /**
   * Highlight the first column (bold)
   * @default false
   */
  showFirstColumn?: boolean;
  /**
   * Highlight the last column (bold)
   * @default false
   */
  showLastColumn?: boolean;
  /**
   * Alternate rows shown with background colour
   * @default false
   */
  showRowStripes?: boolean;
  /**
   * Alternate rows shown with background colour
   * @default false
   */
  showColumnStripes?: boolean;
}

interface TableColumnProperties {
  /**
   * The name of the column, also used in the header
   */
  name: string;
  /**
   * Switches the filter control in the header
   * @default false
   */
  filterButton?: boolean;
  /**
   * Label to describe the totals row (first column)
   * @default 'Total'
   */
  totalsRowLabel?: string;
  /**
   * Name of the totals function
   * @default 'none'
   */
  totalsRowFunction?:
    | 'none'
    | 'average'
    | 'countNums'
    | 'count'
    | 'max'
    | 'min'
    | 'stdDev'
    | 'var'
    | 'sum'
    | 'custom';
  /**
   * Optional formula for custom functions
   */
  totalsRowFormula?: string;

  /**
   * Styles applied to the column
   */
  style?: Partial<FullStyle>;
}

interface TableProperties {
  /**
   * The name of the table
   */
  name: string;
  /**
   * The display name of the table
   */
  displayName?: string;
  /**
   * Top left cell of the table
   */
  ref: string;
  /**
   * Show headers at top of table
   * @default true
   */
  headerRow?: boolean;
  /**
   * Show totals at bottom of table
   * @default false
   */
  totalsRow?: boolean;
  /**
   * Extra style properties
   * @default {}
   */
  style?: TableStyleProperties;
  /**
   * Column definitions
   */
  columns: TableColumnProperties[];
  /**
   * Rows of data
   */
  rows: any[][];
}

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
  views?: LegacyWorkbookView[];
}

export interface ColumnInput {
  /**
   * Header label. Only the first line is used when `header` is an array
   * (multi-row headers are not supported — extra lines are dropped).
   */
  header?: string | string[];
  key?: string;
  width?: number;
  hidden?: boolean;
  style?: Style;
  outlineLevel?: number;
}

export type RowInput = ReadonlyArray<CellValue> | Record<string, CellValue>;

export type SheetTitleInput =
  | string
  | {
      text: string;
      /** Applied to `merge` when given, otherwise to the title cell. */
      style?: Style;
      /** Valid A1 range (e.g. `'A1:B1'`). Invalid ranges throw at build time. */
      merge?: string;
    };

export interface SheetCell {
  value: CellValue;
  style?: Style;
}

export interface SheetRow {
  number: number;
  cells: Record<number, SheetCell>;
  height?: number;
  hidden?: boolean;
  style?: Style;
}

export type NoteValue = string | Comment;
export type WorksheetViewInput = Partial<WorksheetView>;
export type ProtectOptions = Partial<WorksheetProtection> & {spinCount?: number};

export interface MediaImage {
  extension: 'jpeg' | 'png' | 'gif';
  base64?: string;
  filename?: string;
  buffer?: unknown;
}

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

export interface ProtectConfig {
  password?: string;
  options?: ProtectOptions;
}

export interface DefinedNameEntry {
  name: string;
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
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  dataValidations?: Record<string, DataValidation>;
  conditionalFormattings?: ConditionalFormattingOptions[];
  notes?: Record<string, NoteValue>;
  protect?: ProtectConfig;
  sheetProtection?: Record<string, unknown>;
  tables?: TableProperties[];
  images?: SheetImagePlacement[];
}

export interface Workbook {
  meta: WorkbookMeta;
  sheets: SheetModel[];
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
  zip?: unknown;
}

export interface LoadOptions {
  /** Treat input as base64 when data is a string. */
  base64?: boolean;
  /** XML node names to ignore while parsing. */
  ignoreNodes?: string[];
}

export type ReadRowsFormat = 'auto' | 'csv' | 'xlsx';
export type ViewFormat = ReadRowsFormat;

export type ColSlice = {start?: number; end?: number} | number[] | string[];

export interface ViewWorkbookOptions {
  format?: ViewFormat;
  /** Optional name for format sniffing (e.g. `file.name`). */
  filename?: string;
  /** @deprecated Prefer `filename`. */
  name?: string;
  encoding?: string;
}

export interface RowsOptions {
  /** 1-based inclusive start row. Default `1`. */
  start?: number;
  /** 1-based inclusive end row. Default last used row. */
  end?: number;
  cols?: ColSlice;
  range?: string;
  values?: 'string' | 'cell';
  blankrows?: boolean;
  trim?: boolean;
  defval?: string;
}

export interface RecordsOptions extends RowsOptions {
  header?: boolean | number;
}

export interface SheetView {
  readonly name: string;
  readonly id: number;
  readonly index: number;
  readonly protected: boolean;
  rows(options?: RowsOptions & {values?: 'string'}): string[][];
  rows(options: RowsOptions & {values: 'cell'}): CellValue[][];
  rows(options?: RowsOptions): string[][] | CellValue[][];
  records(options?: RecordsOptions): Record<string, string | CellValue>[];
}

export interface WorkbookView {
  readonly format: 'csv' | 'xlsx';
  readonly sheetNames: string[];
  readonly meta: WorkbookMeta;
  sheet(nameOrIndex?: string | number): SheetView;
  toJSON(): Workbook;
}

/** Options for `readRows` (CSV/xlsx → dense `string[][]`). */
export interface ReadRowsOptions extends ViewWorkbookOptions, RowsOptions {
  /** Sheet name or 0-based index. Default `0`. */
  sheet?: string | number;
}

export interface SheetInit {
  title?: SheetTitleInput;
  columns?: ColumnInput[];
  rows?: RowInput[];
  merges?: string[];
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
}

/** Options for `csv.parse` / Node `readCsvFile`. */
export interface CsvParseOptions {
  /** Sheet name used by Node `readCsvFile` (default `"Sheet1"`). Ignored by `csv.parse`. */
  sheetName?: string;
  dateFormats?: string[];
  map?: (datum: string, index?: number) => unknown;
  /** Pass-through to fast-csv parse options. */
  parserOptions?: Record<string, unknown>;
}

/** Options for `csv.stringify` / builder `.csv()` / Node `writeCsvFile`. */
export interface CsvStringifyOptions {
  /** Sheet name to export (default: active/first sheet). */
  sheetName?: string;
  /** Sheet model `id` or 1-based index into `sheets`. */
  sheetId?: number;
  dateFormat?: string;
  dateUTC?: boolean;
  map?: (value: unknown, index?: number) => unknown;
  /** When true (default), emit blank lines for missing row numbers. */
  includeEmptyRows?: boolean;
  /** Pass-through to fast-csv format options. */
  formatterOptions?: Record<string, unknown>;
  /** File encoding for Node `writeCsvFile` only. */
  encoding?: BufferEncoding;
}

export interface SheetBuilder {
  title(title: SheetTitleInput): SheetBuilder;
  row(values: RowInput): SheetBuilder;
  rows(values: RowInput[]): SheetBuilder;
  cell(address: string, value: CellValue, style?: Style): SheetBuilder;
  cells(map: Record<string, CellValue>): SheetBuilder;
  style(range: string, style: Style): SheetBuilder;
  merge(range: string): SheetBuilder;
  columns(cols: ColumnInput[]): SheetBuilder;
  views(views: WorksheetViewInput[]): SheetBuilder;
  pageSetup(setup: Partial<PageSetup>): SheetBuilder;
  headerFooter(hf: Partial<HeaderFooter>): SheetBuilder;
  dataValidation(address: string, rules: DataValidation): SheetBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): SheetBuilder;
  note(address: string, note: NoteValue): SheetBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): SheetBuilder;
  table(table: TableProperties): SheetBuilder;
  image(imageId: number, range: SheetImageRange): SheetBuilder;
}

export interface WorkbookBuilder {
  sheet(name: string, init?: SheetInit | ((s: SheetBuilder) => void)): WorkbookBuilder;
  /** Emit a title row on the active sheet (call `.sheet(name)` first). */
  title(title: SheetTitleInput): WorkbookBuilder;
  row(values: RowInput): WorkbookBuilder;
  rows(values: RowInput[]): WorkbookBuilder;
  cell(address: string, value: CellValue, style?: Style): WorkbookBuilder;
  cells(map: Record<string, CellValue>): WorkbookBuilder;
  style(range: string, style: Style): WorkbookBuilder;
  merge(range: string): WorkbookBuilder;
  columns(cols: ColumnInput[]): WorkbookBuilder;
  props(meta: WorkbookInit): WorkbookBuilder;
  views(views: WorksheetViewInput[]): WorkbookBuilder;
  pageSetup(setup: Partial<PageSetup>): WorkbookBuilder;
  headerFooter(hf: Partial<HeaderFooter>): WorkbookBuilder;
  dataValidation(address: string, rules: DataValidation): WorkbookBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): WorkbookBuilder;
  note(address: string, note: NoteValue): WorkbookBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): WorkbookBuilder;
  table(table: TableProperties): WorkbookBuilder;
  /** Register workbook media; returns image id. */
  image(def: MediaImage): number;
  /** Alias for `image(def)` — clearer name for the register step. */
  addImage(def: MediaImage): number;
  /** Place a registered image on the active sheet. */
  image(imageId: number, range: SheetImageRange): WorkbookBuilder;
  definedName(name: string, refersTo: string): WorkbookBuilder;
  build(): Workbook;
  writeBuffer(opts?: WriteOptions): Promise<Uint8Array>;
  /** Stringify the active (or first) sheet as CSV. */
  csv(opts?: CsvStringifyOptions): Promise<string>;
}

/** Named CSV helpers. */
export const csv: {
  parse(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
  stringify(input: Workbook | WorkbookBuilder, opts?: CsvStringifyOptions): Promise<string>;
};

export function parseCsv(text: string, opts?: CsvParseOptions): Promise<SheetInit>;
export function stringifyCsv(
  input: Workbook | WorkbookBuilder,
  opts?: CsvStringifyOptions,
): Promise<string>;

/**
 * Create a builder, optionally seeded from a `WorkbookInit`, a plain `Workbook`
 * snapshot, or a read-only `WorkbookView`.
 *
 * WARNING: `WorkbookView` (`viewWorkbook`/`readRows`) is values-only (no styles,
 * merges, formulas, hyperlinks, images; dates arrive numeric) — `workbook(view)`
 * re-encodes from values, so never use it for fidelity round-trips. Use `load()`
 * when formatting must survive.
 */
export function workbook(init?: WorkbookInit | Workbook | WorkbookView): WorkbookBuilder;
export function isWorkbookBuilder(value: unknown): value is WorkbookBuilder;
export function writeBuffer(
  input: Workbook | WorkbookBuilder,
  options?: WriteOptions,
): Promise<Uint8Array>;
export function load(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
  options?: LoadOptions,
): Promise<Workbook>;

/** Read-only view over CSV or OOXML (xlsx/xlsm/…). */
export function viewWorkbook(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ViewWorkbookOptions,
): Promise<WorkbookView>;
export function isWorkbookView(value: unknown): value is WorkbookView;

/**
 * CSV or xlsx → dense `string[][]`.
 * Sugar: `(await viewWorkbook(data, opts)).sheet(opts.sheet ?? 0).rows({ values: 'string', ... })`
 */
export function readRows(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ReadRowsOptions,
): Promise<string[][]>;
export function cellToDisplayString(value: CellValue, defval?: string): string;

export const ValueType: typeof LegacyValueType;
export const FormulaType: typeof LegacyFormulaType;
export const RelationshipType: typeof LegacyRelationshipType;
export const DocumentType: typeof LegacyDocumentType;
export const ReadingOrder: typeof LegacyReadingOrder;
export const ErrorValue: typeof LegacyErrorValue;
export const enums: unknown;

// Re-export feature types used by builder methods
export type {
  DataValidation,
  ConditionalFormattingOptions,
  PageSetup,
  HeaderFooter,
  TableProperties,
  Comment,
  WorksheetView,
  WorksheetProtection,
};
