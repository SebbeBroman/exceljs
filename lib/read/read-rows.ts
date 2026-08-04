/**
 * Sugar: CSV/xlsx → dense `string[][]` via {@link viewWorkbook} / light package.
 *
 *   const rows = await readRows(arrayBuffer, { filename: file.name });
 */

import {sniffFormat, toUint8Array} from './format.js';
import {
  viewWorkbook,
  rowsOptionsToLight,
  type RowsOptions,
  type ViewWorkbookOptions,
} from './view.js';
import {openLightPackage, parseLightSheet} from './xlsx-light.js';

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
 *
 * Xlsx uses the light package path directly (no full workbook materialize).
 */
export async function readRows(
  data: ArrayBuffer | Uint8Array | ArrayBufferView | string,
  options?: ReadRowsOptions,
): Promise<string[][]> {
  const opts = options ?? {};
  const {sheet, ...rest} = opts;
  const filename = rest.filename ?? rest.name;
  const binaryOrText =
    typeof data === 'string' ? data : toUint8Array(data as ArrayBuffer | Uint8Array | ArrayBufferView);

  const format = sniffFormat(
    typeof data === 'string' ? data : (binaryOrText as Uint8Array),
    filename,
    rest.format ?? 'auto',
  );

  if (format === 'xlsx') {
    const bytes =
      typeof data === 'string' ? new TextEncoder().encode(data) : (binaryOrText as Uint8Array);
    const pkg = await openLightPackage(bytes);
    const lightOpts = rowsOptionsToLight({
      start: rest.start,
      end: rest.end,
      cols: rest.cols,
      range: rest.range,
      values: 'string',
      blankrows: rest.blankrows,
      trim: rest.trim,
      defval: rest.defval,
    });
    lightOpts.values = 'string';
    return parseLightSheet(pkg, sheet ?? 0, lightOpts).rows as string[][];
  }

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
