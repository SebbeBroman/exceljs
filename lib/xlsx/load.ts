/**
 * Public load API — decode xlsx bytes into a plain Workbook snapshot.
 * Write path stays in write-buffer.ts (no static import of load).
 */

import type {LoadOptions, Workbook} from '../model/types.js';
import DocWorkbook from '../doc/workbook.js';
import {docWorkbookToPlain} from '../compile/doc-to-plain.js';
import type {XlsxReadOptions} from './xlsx.js';

export type {LoadOptions} from '../model/types.js';

/**
 * Load an xlsx workbook from bytes into a plain `{ meta, sheets }` snapshot.
 * Does not export the mutable DocWorkbook graph.
 */
export async function load(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
  opts?: LoadOptions,
): Promise<Workbook> {
  const doc = new DocWorkbook();
  const xlsxOpts: XlsxReadOptions | undefined = opts
    ? {
        base64: opts.base64,
        ignoreNodes: opts.ignoreNodes,
      }
    : undefined;
  await doc.xlsx.load(data, xlsxOpts);
  return docWorkbookToPlain(doc);
}
