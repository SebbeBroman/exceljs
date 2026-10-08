/** Structural XLSX feature types. No mutable document API. */
export const enum PaperSize {
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

export interface WorksheetViewCommon {
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

export interface WorksheetViewNormal {
  /**
   * Controls the view state
   */
  state: 'normal';

  /**
   * Presentation style
   */
  style: 'pageBreakPreview' | 'pageLayout';
}

export interface WorksheetViewFrozen {
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

export interface WorksheetViewSplit {
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

export type WorksheetView = WorksheetViewCommon &
  (WorksheetViewNormal | WorksheetViewFrozen | WorksheetViewSplit);

export interface WorkbookView {
  x: number;
  y: number;
  width: number;
  height: number;
  firstSheet: number;
  activeTab: number;
  visibility: string;
}

export type FillPatterns =
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

export interface FillPattern {
  type: 'pattern';
  pattern: FillPatterns;
  fgColor?: Partial<Color>;
  bgColor?: Partial<Color>;
}

export interface GradientStop {
  position: number;
  color: Partial<Color>;
}

export interface FillGradientAngle {
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

export interface FillGradientPath {
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

export type Fill = FillPattern | FillGradientAngle | FillGradientPath;

export interface Font {
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

export type BorderStyle =
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

export interface Color {
  /**
   * Hex string for alpha-red-green-blue e.g. FF00FF00
   */
  argb: string;

  /**
   * Choose a theme by index
   */
  theme: number;
}

export interface Border {
  style: BorderStyle;
  color: Partial<Color>;
}

export interface BorderDiagonal extends Border {
  up: boolean;
  down: boolean;
}

export interface Borders {
  top: Partial<Border>;
  left: Partial<Border>;
  bottom: Partial<Border>;
  right: Partial<Border>;
  diagonal: Partial<BorderDiagonal>;
}

export interface Margins {
  top: number;
  left: number;
  bottom: number;
  right: number;
  header: number;
  footer: number;
}

export interface Alignment {
  horizontal: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
  wrapText: boolean;
  shrinkToFit: boolean;
  indent: number;
  readingOrder: 'rtl' | 'ltr';
  textRotation: number | 'vertical';
}

export interface Protection {
  locked: boolean;
  hidden: boolean;
}

export interface Style {
  numFmt: string;
  font: Partial<Font>;
  alignment: Partial<Alignment>;
  protection: Partial<Protection>;
  border: Partial<Borders>;
  fill: Fill;
}

export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

export interface DataValidation {
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

export interface CellErrorValue {
  error: '#N/A' | '#REF!' | '#NAME?' | '#DIV/0!' | '#NULL!' | '#VALUE!' | '#NUM!';
}

export interface RichText {
  text: string;
  font?: Partial<Font>;
}

export interface CellRichTextValue {
  richText: RichText[];
}

export interface CellHyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

export interface CellFormulaValue {
  formula: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

export interface CellSharedFormulaValue {
  sharedFormula: string;
  readonly formula?: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

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

export interface CommentMargins {
  insetmode: 'auto' | 'custom';
  inset: number[];
}

export interface CommentProtection {
  locked: 'True' | 'False';
  lockText: 'True' | 'False';
}

export type CommentEditAs = 'twoCells' | 'oneCells' | 'absolute';

export interface Comment {
  texts?: RichText[];
  margins?: Partial<CommentMargins>;
  protection?: Partial<CommentProtection>;
  editAs?: CommentEditAs;
}

export interface PageSetup {
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

export interface HeaderFooter {
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

export type AutoFilter =
  | string
  | {
      from: string | {row: number; column: number};
      to: string | {row: number; column: number};
    };

export interface WorksheetProtection {
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
export interface ImageHyperlinkValue {
  hyperlink: string;
  tooltip?: string;
}

export interface RowBreak {
  id: number;
  max: number;
  min: number;
  man: number;
}

export type WorksheetState = 'visible' | 'hidden' | 'veryHidden';

export type CellIsOperators = 'equal' | 'greaterThan' | 'lessThan' | 'between';

export type ContainsTextOperators =
  | 'containsText'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors';

export type TimePeriodTypes =
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

export type IconSetTypes =
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

export type CfvoTypes =
  | 'percentile'
  | 'percent'
  | 'num'
  | 'min'
  | 'max'
  | 'formula'
  | 'autoMin'
  | 'autoMax';

export interface Cvfo {
  type: CfvoTypes;
  value?: number;
}
export interface ConditionalFormattingBaseRule {
  priority: number;
  style?: Partial<Style>;
}
export interface ExpressionRuleType extends ConditionalFormattingBaseRule {
  type: 'expression';
  formulae?: any[];
}

export interface CellIsRuleType extends ConditionalFormattingBaseRule {
  type: 'cellIs';
  formulae?: any[];
  operator?: CellIsOperators;
}

export interface Top10RuleType extends ConditionalFormattingBaseRule {
  type: 'top10';
  rank: number;
  percent: boolean;
  bottom: boolean;
}

export interface AboveAverageRuleType extends ConditionalFormattingBaseRule {
  type: 'aboveAverage';
  aboveAverage: boolean;
}

export interface ColorScaleRuleType extends ConditionalFormattingBaseRule {
  type: 'colorScale';
  cfvo?: Cvfo[];
  color?: Partial<Color>[];
}

export interface IconSetRuleType extends ConditionalFormattingBaseRule {
  type: 'iconSet';
  showValue?: boolean;
  reverse?: boolean;
  custom?: boolean;
  iconSet?: IconSetTypes;
  cfvo?: Cvfo[];
}

export interface ContainsTextRuleType extends ConditionalFormattingBaseRule {
  type: 'containsText';
  operator?: ContainsTextOperators;
  text?: string;
}

export interface TimePeriodRuleType extends ConditionalFormattingBaseRule {
  type: 'timePeriod';
  timePeriod?: TimePeriodTypes;
}

export interface DataBarRuleType extends ConditionalFormattingBaseRule {
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

export type ConditionalFormattingRule =
  | ExpressionRuleType
  | CellIsRuleType
  | Top10RuleType
  | AboveAverageRuleType
  | ColorScaleRuleType
  | IconSetRuleType
  | ContainsTextRuleType
  | TimePeriodRuleType
  | DataBarRuleType;

export interface ConditionalFormattingOptions {
  ref: string;
  rules: ConditionalFormattingRule[];
}

export interface WorksheetProperties {
  /**
   * Color of the tab
   */
  tabColor: Partial<Color>;

  /**
   * The worksheet column outline level (default: 0)
   */
  outlineLevelCol: number;

  /**
   * The worksheet row outline level (default: 0)
   */
  outlineLevelRow: number;

  /**
   * The outline properties which controls how it will summarize rows and columns
   */
  outlineProperties: {
    summaryBelow: boolean;
    summaryRight: boolean;
  };
  /**
   * Default row height (default: 15)
   */
  defaultRowHeight: number;

  /**
   * Default column width (optional)
   */
  defaultColWidth?: number;

  /**
   * default: 55
   */
  dyDescent: number;
  showGridLines: boolean;
}

export interface WorkbookProperties {
  /**
   * Set workbook dates to 1904 date system
   */
  date1904: boolean;
}

export interface Location {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface DefinedNamesRanges {
  name: string;
  ranges: string[];
}

export type DefinedNamesModel = DefinedNamesRanges[];

export interface TableStyleProperties {
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

export interface TableColumnProperties {
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
  style?: Partial<Style>;
}

export interface TableProperties {
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
