/** Shared value conversion for buffered and streaming XLSX models. */
import Enums from './enums.js';
import type {Style} from './schema.js';
export interface CellValueModel {
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

export interface FormulaValueInput {
  formula?: string;
  sharedFormula?: string;
  shareType?: string;
  ref?: string;
  result?: unknown;
}

export interface HyperlinkValueInput {
  text?: string;
  hyperlink?: string;
  tooltip?: string;
  /** Formula source, when the hyperlink cell was produced from a formula. */
  formula?: string;
  sharedFormula?: string;
  shareType?: string;
  ref?: string;
}

export function getValueType(value: unknown): number {
  if (value === null || value === undefined) {
    return Enums.ValueType.Null;
  }
  if (value instanceof String || typeof value === 'string') {
    return Enums.ValueType.String;
  }
  if (typeof value === 'number') {
    return Enums.ValueType.Number;
  }
  if (typeof value === 'boolean') {
    return Enums.ValueType.Boolean;
  }
  if (value instanceof Date) {
    return Enums.ValueType.Date;
  }
  if (typeof value === 'object' && value !== null) {
    const obj = value as Record<string, unknown>;
    if (obj.text && obj.hyperlink) {
      return Enums.ValueType.Hyperlink;
    }
    if (obj.formula || obj.sharedFormula) {
      return Enums.ValueType.Formula;
    }
    if (obj.richText) {
      return Enums.ValueType.RichText;
    }
    if (obj.sharedString) {
      return Enums.ValueType.SharedString;
    }
    if (obj.error) {
      return Enums.ValueType.Error;
    }
  }
  return undefined as unknown as number;
}

export function valueToModel(address: string, value: unknown, knownType?: number): CellValueModel {
  const type = knownType === undefined ? getValueType(value) : knownType;
  switch (type) {
    case Enums.ValueType.Null:
      return {address, type: Enums.ValueType.Null};
    case Enums.ValueType.Number:
    case Enums.ValueType.String:
    case Enums.ValueType.Date:
    case Enums.ValueType.Boolean:
    case Enums.ValueType.SharedString:
    case Enums.ValueType.Error:
      return {address, type, value};
    case Enums.ValueType.RichText:
      // RichTextValue stores model.type as String (historical)
      return {address, type: Enums.ValueType.String, value};
    case Enums.ValueType.Hyperlink: {
      const v = value as HyperlinkValueInput;
      const model: CellValueModel = {
        address,
        type: Enums.ValueType.Hyperlink,
        text: v.text,
        hyperlink: v.hyperlink,
      };
      if (v.tooltip) model.tooltip = v.tooltip;
      // Keep formula fields so a formula-backed hyperlink writes back intact.
      if (v.formula) model.formula = v.formula;
      if (v.sharedFormula) model.sharedFormula = v.sharedFormula;
      if (v.shareType) model.shareType = v.shareType;
      if (v.ref) model.ref = v.ref;
      return model;
    }
    case Enums.ValueType.Formula: {
      const v = value as FormulaValueInput;
      return {
        address,
        type: Enums.ValueType.Formula,
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
        type: Enums.ValueType.String,
        value: JSON.stringify(value),
        rawValue: value,
      };
  }
}

export function mergeStyles(
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

export function valueFromCellModel(cellModel: {type: number; [key: string]: unknown}): unknown {
  const type = cellModel.type as number;
  switch (type) {
    case Enums.ValueType.Null:
      return null;
    case Enums.ValueType.Number:
    case Enums.ValueType.String:
    case Enums.ValueType.Date:
    case Enums.ValueType.Boolean:
    case Enums.ValueType.Error:
    case Enums.ValueType.SharedString:
    case Enums.ValueType.RichText:
      // JSONValue models carry rawValue (rare on load path)
      if (cellModel.rawValue !== undefined) return cellModel.rawValue;
      return cellModel.value;
    case Enums.ValueType.Hyperlink: {
      const v: {
        text?: unknown;
        hyperlink?: unknown;
        tooltip?: unknown;
        formula?: unknown;
        sharedFormula?: unknown;
        shareType?: unknown;
        ref?: unknown;
      } = {
        text: cellModel.text,
        hyperlink: cellModel.hyperlink,
      };
      if (cellModel.tooltip != null) v.tooltip = cellModel.tooltip;
      // A formula cell reconciled to a hyperlink keeps its `formula` (the
      // reconcile moves result -> text). Carry the formula fields through so
      // `cell.model.formula` survives, matching the direct-model load path.
      if (cellModel.formula != null) v.formula = cellModel.formula;
      if (cellModel.sharedFormula != null) v.sharedFormula = cellModel.sharedFormula;
      if (cellModel.shareType != null) v.shareType = cellModel.shareType;
      if (cellModel.ref != null) v.ref = cellModel.ref;
      return v;
    }
    case Enums.ValueType.Formula: {
      const v: {
        formula?: unknown;
        sharedFormula?: unknown;
        shareType?: unknown;
        ref?: unknown;
        result?: unknown;
      } = {};
      if (cellModel.formula != null) v.formula = cellModel.formula;
      if (cellModel.sharedFormula != null) v.sharedFormula = cellModel.sharedFormula;
      if (cellModel.shareType != null) v.shareType = cellModel.shareType;
      if (cellModel.ref != null) v.ref = cellModel.ref;
      if (cellModel.result !== undefined) v.result = cellModel.result;
      return v;
    }
    default:
      return cellModel.value;
  }
}
