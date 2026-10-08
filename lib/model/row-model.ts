/** Plain row/column encoding shared by buffered and streaming writers. */
import type {ColumnInput, RowInput} from './types.js';
import type {ColumnModel, RowModelData} from './xlsx-model.js';
import {isEqual} from '../utils/object.js';
import colCache from '../utils/col-cache.js';
import {valueToModel} from './cell-model.js';

export function columnsToModel(
  columns: ColumnInput[],
  outlineLevelCol = 0,
): ColumnModel[] | undefined {
  const result: ColumnModel[] = [];
  for (let index = 0; index < columns.length; index++) {
    const column = columns[index];
    const width = column.width ?? 9;
    const style = column.style ?? {};
    if (width === 9 && !column.hidden && !column.outlineLevel && !Object.keys(style).length)
      continue;
    const previous = result[result.length - 1];
    if (
      previous &&
      previous.max === index &&
      previous.width === width &&
      previous.hidden === !!column.hidden &&
      previous.outlineLevel === (column.outlineLevel ?? 0) &&
      isEqual(previous.style, style)
    )
      previous.max++;
    else
      result.push({
        min: index + 1,
        max: index + 1,
        width,
        style,
        isCustomWidth: width !== 9,
        hidden: !!column.hidden,
        outlineLevel: column.outlineLevel ?? 0,
        collapsed: !!column.outlineLevel && column.outlineLevel >= outlineLevelCol,
      });
  }
  return result.length ? result : undefined;
}

export function rowToModel(number: number, values: RowInput, columns: ColumnInput[]): RowModelData {
  const cells: RowModelData['cells'] = [];
  let min = 0;
  let max = 0;
  const add = (col: number, value: unknown) => {
    if (value === undefined) return;
    const cell = valueToModel(colCache.encodeAddress(number, col), value);
    if (columns[col - 1]?.style) cell.style = {...columns[col - 1].style};
    cells.push(cell);
    if (!min) min = col;
    max = col;
  };
  if (Array.isArray(values)) {
    const offset = Object.hasOwn(values, '0') ? 1 : 0;
    values.forEach((value, index) => add(index + offset, value));
  } else {
    columns.forEach((column, index) => {
      if (column.key) add(index + 1, (values as Record<string, unknown>)[column.key]);
    });
  }
  return {number, cells, min, max, style: {}, hidden: false, outlineLevel: 0, collapsed: false};
}
