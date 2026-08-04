/**
 * Row/column slice helpers for SheetView.rows / .records.
 * Indices are **1-based inclusive** (Excel-like).
 */

import colCache from '../utils/col-cache.js';

export type ColSlice =
  | {start?: number; end?: number}
  | number[]
  | string[];

export interface SliceBounds {
  rowStart: number;
  rowEnd: number;
  /** 1-based column indices to include, in output order */
  cols: number[];
}

export function parseA1Range(range: string): {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
} {
  const decoded = colCache.decode(range) as {
    top?: number;
    left?: number;
    bottom?: number;
    right?: number;
    row?: number;
    col?: number;
  };
  if (decoded.top != null && decoded.left != null) {
    return {
      rowStart: decoded.top,
      rowEnd: decoded.bottom as number,
      colStart: decoded.left,
      colEnd: decoded.right as number,
    };
  }
  if (decoded.row != null && decoded.col != null) {
    return {
      rowStart: decoded.row,
      rowEnd: decoded.row,
      colStart: decoded.col,
      colEnd: decoded.col,
    };
  }
  throw new Error(`Invalid range: ${range}`);
}

function colLetterToIndex(letter: string): number {
  const s = letter.trim().toUpperCase();
  if (/^\d+$/.test(s)) return Number(s);
  return colCache.l2n(s);
}

/**
 * Resolve column selection.
 * - `{ start, end }` → contiguous span (1-based)
 * - `number[]` → those columns in order
 * - `string[]` → those letters as discrete columns (`['A','C']` → A and C only)
 * - For a letter span use `{ start: colCache, end }` or `range: 'A:C'` / `'A1:C10'`
 */
export function resolveCols(
  cols: ColSlice | undefined,
  maxCol: number,
  rangeColStart?: number,
  rangeColEnd?: number,
): number[] {
  if (rangeColStart != null && rangeColEnd != null) {
    const out: number[] = [];
    for (let c = rangeColStart; c <= rangeColEnd; c++) out.push(c);
    return out;
  }

  if (cols == null) {
    const out: number[] = [];
    for (let c = 1; c <= maxCol; c++) out.push(c);
    return out;
  }

  if (Array.isArray(cols)) {
    if (cols.length === 0) return [];
    if (typeof cols[0] === 'number') {
      return (cols as number[]).filter(c => c >= 1);
    }
    return (cols as string[]).map(colLetterToIndex).filter(c => c >= 1);
  }

  const start = cols.start ?? 1;
  const end = cols.end ?? maxCol;
  const out: number[] = [];
  for (let c = start; c <= end; c++) out.push(c);
  return out;
}

export function resolveSlice(
  opts: {
    start?: number;
    end?: number;
    cols?: ColSlice;
    range?: string;
  },
  maxRow: number,
  maxCol: number,
): SliceBounds {
  if (opts.range) {
    const r = parseA1Range(opts.range);
    const cols = resolveCols(undefined, maxCol, r.colStart, r.colEnd);
    return {
      rowStart: opts.start ?? r.rowStart,
      rowEnd: opts.end ?? r.rowEnd,
      cols,
    };
  }

  const rowStart = opts.start ?? 1;
  const rowEnd = opts.end ?? maxRow;
  const cols = resolveCols(opts.cols, maxCol);
  return {rowStart, rowEnd, cols};
}

export function sheetExtent(sheet: {
  rows: Array<{number: number; cells: Record<number, unknown>}>;
}): {maxRow: number; maxCol: number} {
  let maxRow = 0;
  let maxCol = 0;
  for (const row of sheet.rows) {
    if (row.number > maxRow) maxRow = row.number;
    for (const colStr of Object.keys(row.cells)) {
      const c = Number(colStr);
      if (c > maxCol) maxCol = c;
    }
  }
  return {maxRow, maxCol};
}
