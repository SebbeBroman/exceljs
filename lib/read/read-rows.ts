/**
 * Sugar: CSV/xlsx → dense `string[][]` via {@link viewWorkbook}.
 *
 *   const rows = await readRows(arrayBuffer, { filename: file.name });
 */

import {viewWorkbook, type RowsOptions, type ViewWorkbookOptions} from './view.js';
export {cellToDisplayString} from './cells.js';
export type {ViewFormat as ReadRowsFormat} from './format.js';

export interface ReadRowsOptions extends ViewWorkbookOptions, RowsOptions {
  /**
   * Sheet name or **0-based** index (xlsx). Default `0`.
   */
  sheet?: string | number;
}

/**
 * Read CSV or xlsx into a dense `string[][]`.
 *
 * Equivalent to:
 *   `(await viewWorkbook(data, opts)).sheet(opts.sheet ?? 0).rows({ ...opts, values: 'string' })`
 */
export async function readRows(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ReadRowsOptions,
): Promise<string[][]> {
  const opts = options ?? {};
  const {sheet, ...rest} = opts;
  const view = await viewWorkbook(data, rest);
  return view.sheet(sheet ?? 0).rows({
    start: rest.start,
    end: rest.end,
    cols: rest.cols,
    range: rest.range,
    values: 'string',
    blankrows: rest.blankrows,
    trim: rest.trim,
    defval: rest.defval,
  });
}
