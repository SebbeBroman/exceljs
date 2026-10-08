/**
 * Public load API — decode xlsx bytes into a plain Workbook snapshot.
 * Write path stays in write-buffer.ts (no static import of load).
 */

import type {LoadOptions, Workbook} from '../model/types.js';
import XlsxReader from './xlsx-reader.js';
import type {XlsxWorkbookModel} from '../compile/ops-to-xlsx-model.js';
import {xlsxModelToPlain} from '../compile/xlsx-model-to-plain.js';
import type {XlsxReadOptions} from './xlsx-reader.js';

export type {LoadOptions} from '../model/types.js';

/**
 * Load an xlsx workbook from bytes into a plain `{ meta, sheets }` snapshot.
 */
export async function load(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
  opts?: LoadOptions,
): Promise<Workbook> {
  const host = {model: undefined as unknown as XlsxWorkbookModel};
  const xlsxOpts: XlsxReadOptions | undefined = opts
    ? {
        base64: opts.base64,
        ignoreNodes: opts.ignoreNodes,
      }
    : undefined;
  await new XlsxReader(host).load(data, xlsxOpts);
  return xlsxModelToPlain(host.model);
}
