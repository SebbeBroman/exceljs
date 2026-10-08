/**
 * Public writeBuffer — accepts a builder or a plain Workbook snapshot.
 * Compiles directly into the existing XLSX encoder model.
 *
 * usedFlags (builder `_used`): when the op log never set styles, default
 * `useStyles: false` so XLSX uses StylesXform.Mock and skips style bookkeeping.
 * Callers can still force styles with `{ useStyles: true }`.
 *
 * Note: the current XLSX encoder still loads the styles module (Mock path);
 * this flag avoids building a real stylesheet, not the import itself.
 */

import type {Workbook, WriteOptions} from '../model/types.js';
import {isWorkbookBuilder, type WorkbookBuilder} from '../builder/workbook-builder.js';
import {compileToXlsxModel, plainToXlsxModel} from '../compile/ops-to-xlsx-model.js';
import XLSX from './xlsx.js';
import {compileToPlainWorkbook} from '../compile/ops-to-model.js';

function toUint8Array(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) {
    const v = data as ArrayBufferView;
    return new Uint8Array(v.buffer, v.byteOffset, v.byteLength);
  }
  throw new Error('writeBuffer: unexpected encoder result type');
}

/**
 * Encode a builder or plain workbook to xlsx bytes.
 */
export async function writeBuffer(
  input: Workbook | WorkbookBuilder,
  options?: WriteOptions,
): Promise<Uint8Array> {
  const model = await (isWorkbookBuilder(input)
    ? compileToXlsxModel(input._ops)
    : plainToXlsxModel(input));
  const opts: WriteOptions = {...options};
  // Skip real style bookkeeping when the builder never applied styles.
  // Plain Workbook snapshots may still carry styles without flags — leave default.
  if (isWorkbookBuilder(input) && !input._used.styles && opts.useStyles === undefined) {
    opts.useStyles = false;
  }
  const raw = await new XLSX({model}).writeBuffer(opts);
  return toUint8Array(raw);
}

/** @internal helper for tests */
export function _plainFromBuilder(builder: WorkbookBuilder): Workbook {
  return compileToPlainWorkbook(builder._ops);
}
