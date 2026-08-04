/**
 * Runtime enum-like objects matching index.d.ts ValueType, FormulaType, etc.
 * Kept as const objects (not TS enums) so values stay plain numbers/strings at runtime.
 */

export const ValueType = {
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
} as const;

export type ValueTypeCode = (typeof ValueType)[keyof typeof ValueType];

export const FormulaType = {
  None: 0,
  Master: 1,
  Shared: 2,
} as const;

export type FormulaTypeCode = (typeof FormulaType)[keyof typeof FormulaType];

export const RelationshipType = {
  None: 0,
  OfficeDocument: 1,
  Worksheet: 2,
  CalcChain: 3,
  SharedStrings: 4,
  Styles: 5,
  Theme: 6,
  Hyperlink: 7,
} as const;

export type RelationshipTypeCode = (typeof RelationshipType)[keyof typeof RelationshipType];

export const DocumentType = {
  Xlsx: 1,
} as const;

export type DocumentTypeCode = (typeof DocumentType)[keyof typeof DocumentType];

export const ReadingOrder = {
  LeftToRight: 1,
  RightToLeft: 2,
} as const;

export type ReadingOrderCode = (typeof ReadingOrder)[keyof typeof ReadingOrder];

export const ErrorValue = {
  NotApplicable: '#N/A',
  Ref: '#REF!',
  Name: '#NAME?',
  DivZero: '#DIV/0!',
  Null: '#NULL!',
  Value: '#VALUE!',
  Num: '#NUM!',
} as const;

export type ErrorValueCode = (typeof ErrorValue)[keyof typeof ErrorValue];

const Enums = {
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
} as const;

export default Enums;
