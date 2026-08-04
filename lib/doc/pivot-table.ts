import {objectFromProps, range, toSortedArray} from '../utils/utils.js';

/** Source worksheet surface used by pivot table builders. */
export interface PivotSourceSheet {
  name: string;
  getRow(rowNumber: number): {values: unknown[]};
  getColumn(columnIndex: number): {values: unknown[]};
}

export interface PivotWorksheet {
  workbook: {pivotTables: unknown[]};
}

export interface PivotTableModelInput {
  sourceSheet: PivotSourceSheet;
  rows: string[];
  columns: string[];
  values: string[];
  metric?: string;
}

export interface PivotCacheField {
  name: string;
  sharedItems: unknown[] | null;
}

export interface PivotTable {
  sourceSheet: PivotSourceSheet;
  rows: number[];
  columns: number[];
  values: number[];
  metric: 'sum';
  cacheFields: PivotCacheField[];
  cacheId: string;
}

// TK(2023-10-10): turn this into a class constructor.

function makePivotTable(worksheet: PivotWorksheet, model: PivotTableModelInput): PivotTable {
  // Example `model`:
  // {
  //   // Source of data: the entire sheet range is taken,
  //   // akin to `worksheet1.getSheetValues()`.
  //   sourceSheet: worksheet1,
  //
  //   // Pivot table fields: values indicate field names;
  //   // they come from the first row in `worksheet1`.
  //   rows: ['A', 'B'],
  //   columns: ['C'],
  //   values: ['E'], // only 1 item possible for now
  //   metric: 'sum', // only 'sum' possible for now
  // }

  validate(worksheet, model);

  const {sourceSheet} = model;
  let {rows, columns, values} = model;

  const cacheFields = makeCacheFields(sourceSheet, [...rows, ...columns]);

  // let {rows, columns, values} use indices instead of names;
  // names can then be accessed via `pivotTable.cacheFields[index].name`.
  // *Note*: Using `reduce` as `Object.fromEntries` requires Node 12+;
  // ExcelJS is >=8.3.0 (as of 2023-10-08).
  const nameToIndex = cacheFields.reduce<Record<string, number>>((result, cacheField, index) => {
    result[cacheField.name] = index;
    return result;
  }, {});
  const rowIndices = rows.map(row => nameToIndex[row]);
  const columnIndices = columns.map(column => nameToIndex[column]);
  const valueIndices = values.map(value => nameToIndex[value]);

  // form pivot table object
  return {
    sourceSheet,
    rows: rowIndices,
    columns: columnIndices,
    values: valueIndices,
    metric: 'sum',
    cacheFields,
    // defined in <pivotTableDefinition> of xl/pivotTables/pivotTable1.xml;
    // also used in xl/workbook.xml
    cacheId: '10',
  };
}

function validate(worksheet: PivotWorksheet, model: PivotTableModelInput): void {
  if (worksheet.workbook.pivotTables.length === 1) {
    throw new Error(
      'A pivot table was already added. At this time, ExcelJS supports at most one pivot table per file.',
    );
  }

  if (model.metric && model.metric !== 'sum') {
    throw new Error('Only the "sum" metric is supported at this time.');
  }

  const headerNames = (model.sourceSheet.getRow(1).values as unknown[]).slice(1) as string[];
  const isInHeaderNames = objectFromProps(headerNames, true) as Record<string, boolean>;
  for (const name of [...model.rows, ...model.columns, ...model.values]) {
    if (!isInHeaderNames[name]) {
      throw new Error(`The header name "${name}" was not found in ${model.sourceSheet.name}.`);
    }
  }

  if (!model.rows.length) {
    throw new Error('No pivot table rows specified.');
  }

  if (!model.columns.length) {
    throw new Error('No pivot table columns specified.');
  }

  if (model.values.length !== 1) {
    throw new Error('Exactly 1 value needs to be specified at this time.');
  }
}

function makeCacheFields(
  worksheet: PivotSourceSheet,
  fieldNamesWithSharedItems: string[],
): PivotCacheField[] {
  // Cache fields are used in pivot tables to reference source data.
  //
  // Example
  // -------
  // Turn
  //
  //  `worksheet` sheet values [
  //    ['A', 'B', 'C', 'D', 'E'],
  //    ['a1', 'b1', 'c1', 4, 5],
  //    ...
  //  ];
  //  fieldNamesWithSharedItems = ['A', 'B', 'C'];
  //
  // into
  //
  //  [
  //    { name: 'A', sharedItems: ['a1', 'a2', 'a3'] },
  //    ...
  //  ]

  const names = worksheet.getRow(1).values as unknown[];
  const nameToHasSharedItems = objectFromProps(fieldNamesWithSharedItems, true) as Record<
    string,
    boolean
  >;

  const aggregate = (columnIndex: number): unknown[] => {
    const columnValues = (worksheet.getColumn(columnIndex).values as unknown[]).splice(2);
    const columnValuesAsSet = new Set(columnValues);
    return toSortedArray(columnValuesAsSet);
  };

  // make result
  const result: PivotCacheField[] = [];
  for (const columnIndex of range(1, names.length)) {
    const name = names[columnIndex] as string;
    const sharedItems = nameToHasSharedItems[name] ? aggregate(columnIndex) : null;
    result.push({name, sharedItems});
  }
  return result;
}

export default {makePivotTable};
export {makePivotTable};
