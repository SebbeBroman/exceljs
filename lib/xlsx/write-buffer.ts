/**
 * Public writeBuffer — accepts a builder or a plain Workbook snapshot.
 * Bridge uses the existing XLSX encoder via an internal DocWorkbook.
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
import {materializeDocWorkbook} from '../compile/ops-to-doc-workbook.js';
import {compileToPlainWorkbook} from '../compile/ops-to-model.js';
import DocWorkbook from '../doc/workbook.js';
import type {BuilderOp} from '../builder/ops.js';
import colCache from '../utils/col-cache.js';

function plainToOps(wb: Workbook): BuilderOp[] {
  const ops: BuilderOp[] = [];
  if (wb.meta && Object.keys(wb.meta).length) {
    ops.push({op: 'meta', meta: wb.meta});
  }
  if (wb.media?.length) {
    for (let i = 0; i < wb.media.length; i++) {
      ops.push({op: 'media', id: i, image: wb.media[i]!});
    }
  }
  if (wb.definedNames?.length) {
    for (const dn of wb.definedNames) {
      ops.push({op: 'definedName', name: dn.name, refersTo: dn.refersTo});
    }
  }
  for (const sheet of wb.sheets) {
    ops.push({op: 'sheet', name: sheet.name});
    if (sheet.columns?.length) {
      ops.push({op: 'columns', sheet: sheet.name, columns: sheet.columns});
    }
    for (const row of sheet.rows) {
      for (const [colStr, cell] of Object.entries(row.cells)) {
        const address = `${colCache.n2l(Number(colStr))}${row.number}`;
        ops.push(
          cell.style
            ? {op: 'cell', sheet: sheet.name, address, value: cell.value, style: cell.style}
            : {op: 'cell', sheet: sheet.name, address, value: cell.value},
        );
      }
    }
    for (const range of sheet.merges ?? []) {
      ops.push({op: 'merge', sheet: sheet.name, range});
    }
    if (sheet.views?.length) {
      ops.push({op: 'views', sheet: sheet.name, views: sheet.views});
    }
    if (sheet.pageSetup && Object.keys(sheet.pageSetup).length) {
      ops.push({op: 'pageSetup', sheet: sheet.name, pageSetup: sheet.pageSetup});
    }
    if (sheet.headerFooter && Object.keys(sheet.headerFooter).length) {
      ops.push({op: 'headerFooter', sheet: sheet.name, headerFooter: sheet.headerFooter});
    }
    if (sheet.dataValidations) {
      for (const [address, rules] of Object.entries(sheet.dataValidations)) {
        if (rules) {
          ops.push({op: 'dataValidation', sheet: sheet.name, address, rules});
        }
      }
    }
    if (sheet.conditionalFormattings?.length) {
      for (const cf of sheet.conditionalFormattings) {
        ops.push({op: 'conditionalFormatting', sheet: sheet.name, cf});
      }
    }
    if (sheet.notes) {
      for (const [address, note] of Object.entries(sheet.notes)) {
        ops.push({op: 'note', sheet: sheet.name, address, note});
      }
    }
    if (sheet.protect) {
      ops.push({
        op: 'protect',
        sheet: sheet.name,
        password: sheet.protect.password,
        options: sheet.protect.options,
      });
    } else if (sheet.sheetProtection) {
      ops.push({op: 'sheetProtection', sheet: sheet.name, model: sheet.sheetProtection});
    }
    if (sheet.tables?.length) {
      for (const table of sheet.tables) {
        ops.push({op: 'table', sheet: sheet.name, table});
      }
    }
    if (sheet.images?.length) {
      for (const img of sheet.images) {
        ops.push({op: 'sheetImage', sheet: sheet.name, imageId: img.imageId, range: img.range});
      }
    }
  }
  return ops;
}

function toDocWorkbook(input: Workbook | WorkbookBuilder): InstanceType<typeof DocWorkbook> {
  if (isWorkbookBuilder(input)) {
    return materializeDocWorkbook(input._ops);
  }
  return materializeDocWorkbook(plainToOps(input));
}

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
  const doc = toDocWorkbook(input);
  const opts: WriteOptions = {...options};
  // Skip real style bookkeeping when the builder never applied styles.
  // Plain Workbook snapshots may still carry styles without flags — leave default.
  if (isWorkbookBuilder(input) && !input._used.styles && opts.useStyles === undefined) {
    opts.useStyles = false;
  }
  const raw = await doc.xlsx.writeBuffer(opts);
  return toUint8Array(raw);
}

/** @internal helper for tests */
export function _plainFromBuilder(builder: WorkbookBuilder): Workbook {
  return compileToPlainWorkbook(builder._ops);
}
