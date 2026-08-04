import {slideFormula} from '../utils/shared-formula.js';
import colCache from '../utils/col-cache.js';
import {escapeHtml} from '../utils/escape-html.js';
import Enums from './enums.js';
import Note from './note.js';
import type {
  Alignment,
  Borders,
  CellErrorValue,
  CellHyperlinkValue,
  CellRichTextValue,
  CellValue as PublicCellValue,
  DataValidation,
  Fill,
  Font,
  Protection,
  Style,
} from '../../index.js';
import type {ValueTypeCode} from './enums.js';
import type NoteClass from './note.js';
import type Column from './column.js';
import type Row from './row.js';

// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

/** Internal value strategy held by a Cell. */
interface CellValueStrategy {
  type: number;
  effectiveType: number;
  value: unknown;
  model: CellValueModel;
  address: string;
  toCsvString(): string | number;
  toString(): string;
  release(): void;
  formula?: string;
  result?: unknown;
  formulaType?: number;
  hyperlink?: string;
  master?: Cell;
  isMergedTo?(master: Cell): boolean;
}

interface CellValueModel {
  address: string;
  type: number;
  value?: unknown;
  style?: Partial<Style>;
  comment?: unknown;
  text?: string;
  hyperlink?: string;
  tooltip?: string;
  master?: string;
  formula?: string;
  sharedFormula?: string;
  shareType?: string;
  ref?: string;
  result?: unknown;
  rawValue?: unknown;
  [key: string]: unknown;
}

interface FormulaValueInput {
  formula?: string;
  sharedFormula?: string;
  shareType?: string;
  ref?: string;
  result?: unknown;
}

interface HyperlinkValueInput {
  text?: string;
  hyperlink?: string;
  tooltip?: string;
}

/** Worksheet surface needed via cell.row */
interface CellWorksheet {
  name: string;
  dataValidations: {
    find(address: string): DataValidation | undefined;
    add(address: string, value: DataValidation | undefined): void;
  };
  findCell(address: string): Cell | undefined;
  workbook: {
    definedNames: {
      getNamesEx(address: {sheetName: string; address: string; row: number; col: number}): string[];
      removeAllNames(address: {
        sheetName: string;
        address: string;
        row: number;
        col: number;
      }): void;
      addEx(
        address: {sheetName: string; address: string; row: number; col: number},
        name: string,
      ): void;
      removeEx(
        address: {sheetName: string; address: string; row: number; col: number},
        name: string,
      ): void;
    };
  };
}

interface CellRow {
  worksheet: CellWorksheet;
  number: number;
  style: Partial<Style>;
}

interface CellColumn {
  number: number;
  letter: string;
  style: Partial<Style>;
}

export interface CellCreateOptions {
  /** Initial cell value (avoids NullValue → typed value thrash). */
  value?: unknown;
  /** Skip address validation when the address is known-good (e.g. from colCache). */
  validateAddress?: boolean;
  /** Pre-merged style; when set, row/column style merge is skipped. */
  style?: Partial<Style> & Record<string, unknown>;
}

class Cell {
  static Types = Enums.ValueType;

  _row!: CellRow;
  _column!: CellColumn;
  _address!: string;
  _value!: CellValueStrategy;
  style!: Partial<Style> & Record<string, unknown>;
  _mergeCount: number;
  _comment?: NoteClass | undefined;

  constructor(
    row: CellRow | Row,
    column: CellColumn | Column,
    address: string,
    options?: CellCreateOptions,
  ) {
    if (!row || !column) {
      throw new Error('A Cell needs a Row');
    }

    this._row = row as CellRow;
    this._column = column as CellColumn;

    if (options?.validateAddress !== false) {
      colCache.validateAddress(address);
    }
    this._address = address;

    if (options && Object.prototype.hasOwnProperty.call(options, 'value')) {
      this._value = Value.create(Value.getType(options.value), this, options.value);
    } else {
      this._value = Value.create(Cell.Types.Null, this);
    }

    if (options?.style) {
      this.style = options.style;
    } else {
      this.style = this._mergeStyle(row.style, column.style, {});
    }

    this._mergeCount = 0;
  }

  /** Infer ValueType for a JS value (used by row compact storage). */
  static getValueType(value: unknown): number {
    return Value.getType(value);
  }

  /**
   * Build the wire model for a cell value without allocating a Cell / Value strategy.
   * Shape matches Value*.model used by Cell.model / xforms.
   * @param knownType optional ValueType (or -1 for JSON) to skip re-inference
   */
  static valueToModel(address: string, value: unknown, knownType?: number): CellValueModel {
    const type = knownType === undefined ? Value.getType(value) : knownType;
    switch (type) {
      case Cell.Types.Null:
        return {address, type: Cell.Types.Null};
      case Cell.Types.Number:
      case Cell.Types.String:
      case Cell.Types.Date:
      case Cell.Types.Boolean:
      case Cell.Types.SharedString:
      case Cell.Types.Error:
        return {address, type, value};
      case Cell.Types.RichText:
        // RichTextValue stores model.type as String (historical)
        return {address, type: Cell.Types.String, value};
      case Cell.Types.Hyperlink: {
        const v = value as HyperlinkValueInput;
        const model: CellValueModel = {
          address,
          type: Cell.Types.Hyperlink,
          text: v.text,
          hyperlink: v.hyperlink,
        };
        if (v.tooltip) model.tooltip = v.tooltip;
        return model;
      }
      case Cell.Types.Formula: {
        const v = value as FormulaValueInput;
        return {
          address,
          type: Cell.Types.Formula,
          shareType: v.shareType,
          ref: v.ref,
          formula: v.formula,
          sharedFormula: v.sharedFormula,
          result: v.result,
        };
      }
      case -1:
      default:
        // plain objects → JSON string (JSONValue); knownType -1 skips re-detect
        return {
          address,
          type: Cell.Types.String,
          value: JSON.stringify(value),
          rawValue: value,
        };
    }
  }

  /** Merge row/column styles into a plain object (shared with compact cell path). */
  static mergeStyles(
    rowStyle: Partial<Style> | undefined | null,
    colStyle: Partial<Style> | undefined | null,
    style: Partial<Style> & Record<string, unknown> = {},
  ): Partial<Style> & Record<string, unknown> {
    const numFmt = (rowStyle && rowStyle.numFmt) || (colStyle && colStyle.numFmt);
    if (numFmt) style.numFmt = numFmt;

    const font = (rowStyle && rowStyle.font) || (colStyle && colStyle.font);
    if (font) style.font = font;

    const alignment = (rowStyle && rowStyle.alignment) || (colStyle && colStyle.alignment);
    if (alignment) style.alignment = alignment;

    const border = (rowStyle && rowStyle.border) || (colStyle && colStyle.border);
    if (border) style.border = border;

    const fill = (rowStyle && rowStyle.fill) || (colStyle && colStyle.fill);
    if (fill) style.fill = fill;

    const protection = (rowStyle && rowStyle.protection) || (colStyle && colStyle.protection);
    if (protection) style.protection = protection;

    return style;
  }

  get worksheet(): CellWorksheet {
    return this._row.worksheet;
  }

  get workbook(): CellWorksheet['workbook'] {
    return this._row.worksheet.workbook;
  }

  // help GC by removing cyclic (and other) references
  destroy(): void {
    delete (this as Partial<Cell>).style;
    delete (this as Partial<Cell>)._value;
    delete (this as Partial<Cell>)._row;
    delete (this as Partial<Cell>)._column;
    delete (this as Partial<Cell>)._address;
  }

  // =========================================================================
  // Styles stuff
  get numFmt(): string | undefined {
    return this.style.numFmt as string | undefined;
  }

  set numFmt(value: string | undefined) {
    this.style.numFmt = value;
  }

  get font(): Font | undefined {
    return this.style.font as Font | undefined;
  }

  set font(value: Font | undefined) {
    this.style.font = value;
  }

  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment as Partial<Alignment> | undefined;
  }

  set alignment(value: Partial<Alignment> | undefined) {
    this.style.alignment = value;
  }

  get border(): Partial<Borders> | undefined {
    return this.style.border as Partial<Borders> | undefined;
  }

  set border(value: Partial<Borders> | undefined) {
    this.style.border = value;
  }

  get fill(): Fill | undefined {
    return this.style.fill as Fill | undefined;
  }

  set fill(value: Fill | undefined) {
    this.style.fill = value;
  }

  get protection(): Partial<Protection> | undefined {
    return this.style.protection as Partial<Protection> | undefined;
  }

  set protection(value: Partial<Protection> | undefined) {
    this.style.protection = value;
  }

  _mergeStyle(
    rowStyle: Partial<Style> | undefined | null,
    colStyle: Partial<Style> | undefined | null,
    style: Partial<Style> & Record<string, unknown>,
  ): Partial<Style> & Record<string, unknown> {
    return Cell.mergeStyles(rowStyle, colStyle, style);
  }

  // =========================================================================
  // return the address for this cell
  get address(): string {
    return this._address;
  }

  get row(): number {
    return this._row.number;
  }

  get col(): number {
    return this._column.number;
  }

  get $col$row(): string {
    return `$${this._column.letter}$${this.row}`;
  }

  // =========================================================================
  // Value stuff

  get type(): number {
    return this._value.type;
  }

  get effectiveType(): number {
    return this._value.effectiveType;
  }

  toCsvString(): string | number {
    return this._value.toCsvString();
  }

  // =========================================================================
  // Merge stuff

  addMergeRef(): void {
    this._mergeCount++;
  }

  releaseMergeRef(): void {
    this._mergeCount--;
  }

  get isMerged(): boolean {
    return this._mergeCount > 0 || this.type === Cell.Types.Merge;
  }

  merge(master: Cell, ignoreStyle?: boolean): void {
    this._value.release();
    this._value = Value.create(Cell.Types.Merge, this, master);
    if (!ignoreStyle) {
      this.style = master.style;
    }
  }

  unmerge(): void {
    if (this.type === Cell.Types.Merge) {
      this._value.release();
      this._value = Value.create(Cell.Types.Null, this);
      this.style = this._mergeStyle(this._row.style, this._column.style, {});
    }
  }

  isMergedTo(master: Cell): boolean {
    if (this._value.type !== Cell.Types.Merge) return false;
    return this._value.isMergedTo!(master);
  }

  get master(): Cell {
    if (this.type === Cell.Types.Merge) {
      return this._value.master!;
    }
    return this; // an unmerged cell is its own master
  }

  get isHyperlink(): boolean {
    return this._value.type === Cell.Types.Hyperlink;
  }

  get hyperlink(): string | undefined {
    return this._value.hyperlink;
  }

  // return the value
  get value(): PublicCellValue {
    return this._value.value as PublicCellValue;
  }

  // set the value - can be number, string or raw
  set value(v: PublicCellValue | unknown) {
    // special case - merge cells set their master's value
    if (this.type === Cell.Types.Merge) {
      this._value.master!.value = v as PublicCellValue;
      return;
    }

    this._value.release();

    // assign value
    this._value = Value.create(Value.getType(v), this, v);
  }

  get note(): unknown {
    return this._comment && this._comment.note;
  }

  set note(note: unknown) {
    this._comment = new Note(note as string | import('./note.js').NoteInput);
  }

  get text(): string {
    return this._value.toString();
  }

  get html(): string {
    return escapeHtml(this.text);
  }

  toString(): string {
    return this.text;
  }

  _upgradeToHyperlink(hyperlink: string): void {
    // if this cell is a string, turn it into a Hyperlink
    if (this.type === Cell.Types.String) {
      this._value = Value.create(Cell.Types.Hyperlink, this, {
        text: this._value.value,
        hyperlink,
      });
    }
  }

  // =========================================================================
  // Formula stuff
  get formula(): string | undefined {
    return this._value.formula;
  }

  get result(): unknown {
    return this._value.result;
  }

  get formulaType(): number | undefined {
    return this._value.formulaType;
  }

  // =========================================================================
  // Name stuff
  get fullAddress(): {sheetName: string; address: string; row: number; col: number} {
    const {worksheet} = this._row;
    return {
      sheetName: worksheet.name,
      address: this.address,
      row: this.row,
      col: this.col,
    };
  }

  get name(): string | undefined {
    return this.names[0];
  }

  set name(value: string) {
    this.names = [value];
  }

  get names(): string[] {
    return this.workbook.definedNames.getNamesEx(this.fullAddress);
  }

  set names(value: string[]) {
    const {definedNames} = this.workbook;
    definedNames.removeAllNames(this.fullAddress);
    value.forEach(name => {
      definedNames.addEx(this.fullAddress, name);
    });
  }

  addName(name: string): void {
    this.workbook.definedNames.addEx(this.fullAddress, name);
  }

  removeName(name: string): void {
    this.workbook.definedNames.removeEx(this.fullAddress, name);
  }

  removeAllNames(): void {
    this.workbook.definedNames.removeAllNames(this.fullAddress);
  }

  // =========================================================================
  // Data Validation stuff
  get _dataValidations(): CellWorksheet['dataValidations'] {
    return this.worksheet.dataValidations;
  }

  get dataValidation(): DataValidation | undefined {
    return this._dataValidations.find(this.address);
  }

  set dataValidation(value: DataValidation | undefined) {
    this._dataValidations.add(this.address, value);
  }

  // =========================================================================
  // Model stuff

  get model(): CellValueModel {
    const {model} = this._value;
    model.style = this.style;
    if (this._comment) {
      model.comment = this._comment.model;
    }
    return model;
  }

  set model(value: CellValueModel) {
    this._value.release();
    this._value = Value.create(value.type, this);
    this._value.model = value;

    if (value.comment) {
      switch ((value.comment as {type?: string}).type) {
        case 'note':
          this._comment = Note.fromModel(value.comment as Parameters<typeof Note.fromModel>[0]);
          break;
      }
    }

    if (value.style) {
      this.style = value.style as Partial<Style> & Record<string, unknown>;
    } else {
      this.style = {};
    }
  }
}

// =============================================================================
// Internal Value Types

class NullValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Null,
    };
  }

  get value(): null {
    return null;
  }

  set value(_value: unknown) {
    // nothing to do
  }

  get type(): number {
    return Cell.Types.Null;
  }

  get effectiveType(): number {
    return Cell.Types.Null;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return '';
  }

  release(): void {}

  toString(): string {
    return '';
  }
}

class NumberValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: number) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Number,
      value,
    };
  }

  get value(): number {
    return this.model.value as number;
  }

  set value(value: number) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.Number;
  }

  get effectiveType(): number {
    return Cell.Types.Number;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return (this.model.value as number).toString();
  }

  release(): void {}

  toString(): string {
    return (this.model.value as number).toString();
  }
}

class StringValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: string) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value,
    };
  }

  get value(): string {
    return this.model.value as string;
  }

  set value(value: string) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.String;
  }

  get effectiveType(): number {
    return Cell.Types.String;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return `"${(this.model.value as string).replace(/"/g, '""')}"`;
  }

  release(): void {}

  toString(): string {
    return this.model.value as string;
  }
}

class RichTextValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: CellRichTextValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value,
    };
  }

  get value(): CellRichTextValue {
    return this.model.value as CellRichTextValue;
  }

  set value(value: CellRichTextValue) {
    this.model.value = value;
  }

  toString(): string {
    return (this.model.value as CellRichTextValue).richText.map(t => t.text).join('');
  }

  get type(): number {
    return Cell.Types.RichText;
  }

  get effectiveType(): number {
    return Cell.Types.RichText;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    // Original references this.text (no getter on RichTextValue — same runtime shape)
    return `"${(this as unknown as {text: string}).text.replace(/"/g, '""')}"`;
  }

  release(): void {}
}

class DateValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: Date) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Date,
      value,
    };
  }

  get value(): Date {
    return this.model.value as Date;
  }

  set value(value: Date) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.Date;
  }

  get effectiveType(): number {
    return Cell.Types.Date;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return (this.model.value as Date).toISOString();
  }

  release(): void {}

  toString(): string {
    return (this.model.value as Date).toString();
  }
}

class HyperlinkValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: HyperlinkValueInput) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Hyperlink,
      text: value ? value.text : undefined,
      hyperlink: value ? value.hyperlink : undefined,
    };
    if (value && value.tooltip) {
      this.model.tooltip = value.tooltip;
    }
  }

  get value(): CellHyperlinkValue {
    const v: CellHyperlinkValue = {
      text: this.model.text as string,
      hyperlink: this.model.hyperlink as string,
    };
    if (this.model.tooltip) {
      (v as CellHyperlinkValue & {tooltip?: string}).tooltip = this.model.tooltip as string;
    }
    return v;
  }

  set value(value: HyperlinkValueInput) {
    this.model = {
      text: value.text,
      hyperlink: value.hyperlink,
    } as CellValueModel;
    if (value.tooltip) {
      this.model.tooltip = value.tooltip;
    }
  }

  get text(): string | undefined {
    return this.model.text as string | undefined;
  }

  set text(value: string | undefined) {
    this.model.text = value;
  }

  get hyperlink(): string | undefined {
    return this.model.hyperlink as string | undefined;
  }

  set hyperlink(value: string | undefined) {
    this.model.hyperlink = value;
  }

  get type(): number {
    return Cell.Types.Hyperlink;
  }

  get effectiveType(): number {
    return Cell.Types.Hyperlink;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.hyperlink as string;
  }

  release(): void {}

  toString(): string {
    return this.model.text as string;
  }
}

class MergeValue implements CellValueStrategy {
  model: CellValueModel;
  _master!: Cell;

  constructor(cell: Cell, master?: Cell) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Merge,
      master: master ? master.address : undefined,
    };
    this._master = master!;
    if (master) {
      master.addMergeRef();
    }
  }

  get value(): PublicCellValue {
    return this._master.value;
  }

  set value(value: PublicCellValue | Cell) {
    if (value instanceof Cell) {
      if (this._master) {
        this._master.releaseMergeRef();
      }
      value.addMergeRef();
      this._master = value;
    } else {
      this._master.value = value;
    }
  }

  isMergedTo(master: Cell): boolean {
    return master === this._master;
  }

  get master(): Cell {
    return this._master;
  }

  get type(): number {
    return Cell.Types.Merge;
  }

  get effectiveType(): number {
    return this._master.effectiveType;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return '';
  }

  release(): void {
    this._master.releaseMergeRef();
  }

  toString(): string {
    return (this.value as {toString(): string}).toString();
  }
}

class FormulaValue implements CellValueStrategy {
  cell: Cell;
  model: CellValueModel;
  _translatedFormula?: string | null;

  constructor(cell: Cell, value?: FormulaValueInput) {
    this.cell = cell;

    this.model = {
      address: cell.address,
      type: Cell.Types.Formula,
      shareType: value ? value.shareType : undefined,
      ref: value ? value.ref : undefined,
      formula: value ? value.formula : undefined,
      sharedFormula: value ? value.sharedFormula : undefined,
      result: value ? value.result : undefined,
    };
  }

  _copyModel(model: FormulaValueInput): FormulaValueInput {
    const copy: FormulaValueInput = {};
    const cp = (name: keyof FormulaValueInput) => {
      const value = model[name];
      if (value) {
        (copy as Record<string, unknown>)[name] = value;
      }
    };
    cp('formula');
    cp('result');
    cp('ref');
    cp('shareType');
    cp('sharedFormula');
    return copy;
  }

  get value(): FormulaValueInput {
    return this._copyModel(this.model);
  }

  set value(value: FormulaValueInput) {
    this.model = this._copyModel(value) as CellValueModel;
  }

  validate(value: unknown): void {
    switch (Value.getType(value)) {
      case Cell.Types.Null:
      case Cell.Types.String:
      case Cell.Types.Number:
      case Cell.Types.Date:
        break;
      case Cell.Types.Hyperlink:
      case Cell.Types.Formula:
      default:
        throw new Error('Cannot process that type of result value');
    }
  }

  get dependencies(): {ranges: RegExpMatchArray | null; cells: RegExpMatchArray | null} {
    // find all the ranges and cells mentioned in the formula
    const formula = this.formula ?? '';
    const ranges = formula.match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g);
    const cells = formula
      .replace(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g, '')
      .match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g);
    return {
      ranges,
      cells,
    };
  }

  get formula(): string | undefined {
    return (this.model.formula as string | undefined) || this._getTranslatedFormula() || undefined;
  }

  set formula(value: string | undefined) {
    this.model.formula = value;
  }

  get formulaType(): number {
    if (this.model.formula) {
      return Enums.FormulaType.Master;
    }
    if (this.model.sharedFormula) {
      return Enums.FormulaType.Shared;
    }
    return Enums.FormulaType.None;
  }

  get result(): unknown {
    return this.model.result;
  }

  set result(value: unknown) {
    this.model.result = value;
  }

  get type(): number {
    return Cell.Types.Formula;
  }

  get effectiveType(): number {
    const v = this.model.result as unknown;
    if (v === null || v === undefined) {
      return Enums.ValueType.Null;
    }
    if (v instanceof String || typeof v === 'string') {
      return Enums.ValueType.String;
    }
    if (typeof v === 'number') {
      return Enums.ValueType.Number;
    }
    if (v instanceof Date) {
      return Enums.ValueType.Date;
    }
    if (typeof v === 'object' && v !== null && 'text' in v && 'hyperlink' in v) {
      return Enums.ValueType.Hyperlink;
    }
    if (typeof v === 'object' && v !== null && 'formula' in v) {
      return Enums.ValueType.Formula;
    }

    return Enums.ValueType.Null;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  _getTranslatedFormula(): string | null | undefined {
    if (!this._translatedFormula && this.model.sharedFormula) {
      const {worksheet} = this.cell;
      const master = worksheet.findCell(this.model.sharedFormula as string);
      this._translatedFormula =
        master &&
        slideFormula(master.formula ?? '', master.address, this.model.address);
    }
    return this._translatedFormula;
  }

  toCsvString(): string {
    return `${this.model.result || ''}`;
  }

  release(): void {}

  toString(): string {
    return this.model.result ? String(this.model.result) : '';
  }
}

class SharedStringValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: unknown) {
    this.model = {
      address: cell.address,
      type: Cell.Types.SharedString,
      value,
    };
  }

  get value(): unknown {
    return this.model.value;
  }

  set value(value: unknown) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.SharedString;
  }

  get effectiveType(): number {
    return Cell.Types.SharedString;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return String(this.model.value);
  }

  release(): void {}

  toString(): string {
    return String(this.model.value);
  }
}

class BooleanValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: boolean) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Boolean,
      value,
    };
  }

  get value(): boolean {
    return this.model.value as boolean;
  }

  set value(value: boolean) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.Boolean;
  }

  get effectiveType(): number {
    return Cell.Types.Boolean;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): number {
    return this.model.value ? 1 : 0;
  }

  release(): void {}

  toString(): string {
    return String(this.model.value);
  }
}

class ErrorValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: CellErrorValue) {
    this.model = {
      address: cell.address,
      type: Cell.Types.Error,
      value,
    };
  }

  get value(): CellErrorValue {
    return this.model.value as CellErrorValue;
  }

  set value(value: CellErrorValue) {
    this.model.value = value;
  }

  get type(): number {
    return Cell.Types.Error;
  }

  get effectiveType(): number {
    return Cell.Types.Error;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.toString();
  }

  release(): void {}

  toString(): string {
    return (this.model.value as CellErrorValue).error.toString();
  }
}

class JSONValue implements CellValueStrategy {
  model: CellValueModel;

  constructor(cell: Cell, value?: unknown) {
    this.model = {
      address: cell.address,
      type: Cell.Types.String,
      value: JSON.stringify(value),
      rawValue: value,
    };
  }

  get value(): unknown {
    return this.model.rawValue;
  }

  set value(value: unknown) {
    this.model.rawValue = value;
    this.model.value = JSON.stringify(value);
  }

  get type(): number {
    return Cell.Types.String;
  }

  get effectiveType(): number {
    return Cell.Types.String;
  }

  get address(): string {
    return this.model.address;
  }

  set address(value: string) {
    this.model.address = value;
  }

  toCsvString(): string {
    return this.model.value as string;
  }

  release(): void {}

  toString(): string {
    return this.model.value as string;
  }
}

// Original code referenced Cell.Types.JSON which is not on ValueType (undefined at runtime).
// Preserve that behaviour for plain-object values.
const JSON_TYPE = undefined as unknown as number;

type ValueConstructor = new (cell: Cell, value?: unknown) => CellValueStrategy;

// Value is a place to hold common static Value type functions
const Value = {
  getType(value: unknown): number {
    if (value === null || value === undefined) {
      return Cell.Types.Null;
    }
    if (value instanceof String || typeof value === 'string') {
      return Cell.Types.String;
    }
    if (typeof value === 'number') {
      return Cell.Types.Number;
    }
    if (typeof value === 'boolean') {
      return Cell.Types.Boolean;
    }
    if (value instanceof Date) {
      return Cell.Types.Date;
    }
    if (typeof value === 'object' && value !== null) {
      const obj = value as Record<string, unknown>;
      if (obj.text && obj.hyperlink) {
        return Cell.Types.Hyperlink;
      }
      if (obj.formula || obj.sharedFormula) {
        return Cell.Types.Formula;
      }
      if (obj.richText) {
        return Cell.Types.RichText;
      }
      if (obj.sharedString) {
        return Cell.Types.SharedString;
      }
      if (obj.error) {
        return Cell.Types.Error;
      }
    }
    return JSON_TYPE;
  },

  // map valueType to constructor (array + sparse keys; JSON_TYPE is undefined at runtime)
  types: (() => {
    const map: Record<string | number, ValueConstructor> = {};
    (
      [
        {t: Cell.Types.Null, f: NullValue},
        {t: Cell.Types.Number, f: NumberValue},
        {t: Cell.Types.String, f: StringValue},
        {t: Cell.Types.Date, f: DateValue},
        {t: Cell.Types.Hyperlink, f: HyperlinkValue},
        {t: Cell.Types.Formula, f: FormulaValue},
        {t: Cell.Types.Merge, f: MergeValue},
        {t: JSON_TYPE, f: JSONValue},
        {t: Cell.Types.SharedString, f: SharedStringValue},
        {t: Cell.Types.RichText, f: RichTextValue},
        {t: Cell.Types.Boolean, f: BooleanValue},
        {t: Cell.Types.Error, f: ErrorValue},
      ] as Array<{t: number; f: ValueConstructor}>
    ).forEach(entry => {
      map[entry.t as number] = entry.f;
    });
    // Original reduced into an array; preserve array-like object for index access
    return map as unknown as ValueConstructor[] & Record<number | string, ValueConstructor>;
  })(),

  create(type: number, cell: Cell, value?: unknown): CellValueStrategy {
    const T = this.types[type as number];
    if (!T) {
      throw new Error(`Could not create Value of type ${type}`);
    }
    return new T(cell, value);
  },
};

export default Cell;
export {Cell};
export type {ValueTypeCode};
