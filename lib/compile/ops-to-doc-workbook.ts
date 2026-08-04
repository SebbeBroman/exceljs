/**
 * Phase 1–5 bridge: apply builder ops onto the existing mutable Workbook
 * so we can reuse XLSX.writeBuffer without rewriting encoders.
 * Not part of the public API.
 *
 * Dense fast path: when the op log is rectangular-only (sheet/meta/columns/row/rows),
 * bulk-append via `addRows` and skip random cell/style work.
 */

import DocWorkbook from '../doc/workbook.js';
import type {BuilderOp} from '../builder/ops.js';
import type {CellValue, ColumnInput, ProtectOptions, RowInput, Style} from '../model/types.js';
import colCache from '../utils/col-cache.js';
import Encryptor from '../utils/encryptor.js';
import {isDenseRectangularOps, optimizeOps} from './ops-to-model.js';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyWs = any;

function applyStyleToRange(ws: AnyWs, range: string, style: Style): void {
  const decoded = colCache.decode(range) as {
    top?: number;
    left?: number;
    bottom?: number;
    right?: number;
    row?: number;
    col?: number;
    address?: string;
  };

  const assign = (cell: AnyWs) => {
    if (style.font) cell.font = {...cell.font, ...style.font};
    if (style.alignment) cell.alignment = {...cell.alignment, ...style.alignment};
    if (style.border) cell.border = {...cell.border, ...style.border};
    if (style.fill) cell.fill = style.fill;
    if (style.numFmt) cell.numFmt = style.numFmt;
    if (style.protection) cell.protection = {...cell.protection, ...style.protection};
  };

  if (decoded.top != null && decoded.left != null) {
    for (let r = decoded.top; r <= (decoded.bottom as number); r++) {
      for (let c = decoded.left; c <= (decoded.right as number); c++) {
        assign(ws.getCell(r, c));
      }
    }
    return;
  }
  if (decoded.row != null && decoded.col != null) {
    assign(ws.getCell(decoded.row, decoded.col));
    return;
  }
  // row-only style like "1" is not valid A1; support via getRow if numeric string
  if (/^\d+$/.test(range)) {
    const row = ws.getRow(Number(range));
    row.font = style.font ? {...row.font, ...style.font} : row.font;
    if (style.alignment) row.alignment = {...row.alignment, ...style.alignment};
    if (style.border) row.border = {...row.border, ...style.border};
    if (style.fill) row.fill = style.fill;
    if (style.numFmt) row.numFmt = style.numFmt;
  }
}

/** Append a single row without copying array/object inputs. */
function setRowValues(ws: AnyWs, values: RowInput): void {
  // Row.values already materializes compact cells from the input; no need to clone.
  ws.addRow(values as CellValue[] | Record<string, CellValue>);
}

/**
 * Dense bulk path: append many rows in one call.
 * Prefer a single addRows when all entries are arrays (common `.rows([[…],…])` case).
 */
function appendRowsDense(ws: AnyWs, values: readonly RowInput[]): void {
  if (values.length === 0) return;
  let allArrays = true;
  for (let i = 0; i < values.length; i++) {
    if (!Array.isArray(values[i])) {
      allArrays = false;
      break;
    }
  }
  if (allArrays) {
    // One batch; worksheet still uses compact cells per row (no full Cell graph).
    ws.addRows(values as CellValue[][]);
    return;
  }
  for (let i = 0; i < values.length; i++) {
    setRowValues(ws, values[i]!);
  }
}

function applyColumns(ws: AnyWs, columns: ColumnInput[]): void {
  ws.columns = columns.map(c => ({
    header: c.header,
    key: c.key,
    width: c.width,
    hidden: c.hidden,
    style: c.style,
    outlineLevel: c.outlineLevel,
  }));
}

function applyMeta(wb: InstanceType<typeof DocWorkbook>, m: BuilderOp & {op: 'meta'}): void {
  const meta = m.meta;
  if (meta.creator != null) wb.creator = meta.creator;
  if (meta.lastModifiedBy != null) wb.lastModifiedBy = meta.lastModifiedBy;
  if (meta.created != null) wb.created = meta.created;
  if (meta.modified != null) wb.modified = meta.modified;
  if (meta.company != null) wb.company = meta.company;
  if (meta.manager != null) wb.manager = meta.manager;
  if (meta.title != null) wb.title = meta.title;
  if (meta.subject != null) wb.subject = meta.subject;
  if (meta.keywords != null) wb.keywords = meta.keywords;
  if (meta.category != null) wb.category = meta.category;
  if (meta.description != null) wb.description = meta.description;
  if (meta.language != null) wb.language = meta.language;
  if (meta.revision != null) wb.revision = meta.revision;
  if (meta.contentStatus != null) wb.contentStatus = meta.contentStatus;
  if (meta.properties) wb.properties = {...wb.properties, ...meta.properties};
  if (meta.views) wb.views = meta.views;
}

/**
 * Apply deferred sheet protection synchronously (same hash as Worksheet.protect).
 * Keeps the builder chain sync; hashing happens once at materialize.
 */
function applyProtect(ws: AnyWs, password?: string, options?: ProtectOptions): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheetProtection: Record<string, any> = {sheet: true};
  let spinCount: number | undefined;
  if (options && 'spinCount' in options && options.spinCount != null) {
    spinCount = Number.isFinite(options.spinCount)
      ? Math.round(Math.max(0, options.spinCount))
      : 100000;
  }
  if (password) {
    sheetProtection.algorithmName = 'SHA-512';
    sheetProtection.saltValue = Encryptor.randomBytesBase64(16);
    sheetProtection.spinCount = spinCount ?? 100000;
    sheetProtection.hashValue = Encryptor.convertPasswordToHash(
      password,
      'SHA512',
      sheetProtection.saltValue,
      sheetProtection.spinCount,
    );
  }
  if (options) {
    Object.assign(sheetProtection, options);
    if (!password && 'spinCount' in options) {
      delete sheetProtection.spinCount;
    }
  }
  ws.sheetProtection = sheetProtection;
}

function applyDefinedName(wb: InstanceType<typeof DocWorkbook>, name: string, refersTo: string): void {
  // `refersTo` is a sheet-qualified range (Sheet1!$A$1) or similar.
  // DefinedNames.add expects locStr then name.
  try {
    wb.definedNames.add(refersTo, name);
  } catch {
    // Fallback: force model merge for non-standard ranges
    const existing = wb.definedNames.model || [];
    wb.definedNames.model = [...existing.filter(d => d.name !== name), {name, ranges: [refersTo]}];
  }
}

/**
 * Materialize builder ops into an internal DocWorkbook for XLSX encoding.
 * Uses a dense bulk path when the op log is rectangular-only.
 */
export function materializeDocWorkbook(ops: BuilderOp[]): InstanceType<typeof DocWorkbook> {
  const optimized = optimizeOps(ops);
  const dense = isDenseRectangularOps(optimized);
  const wb = new DocWorkbook();
  const sheets = new Map<string, AnyWs>();
  /** media id (from builder) → actual workbook media index */
  const mediaIdMap = new Map<number, number>();

  const ensureSheet = (name: string): AnyWs => {
    let ws = sheets.get(name);
    if (!ws) {
      ws = wb.addWorksheet(name);
      sheets.set(name, ws);
    }
    return ws;
  };

  if (dense) {
    for (const op of optimized) {
      switch (op.op) {
        case 'meta':
          applyMeta(wb, op);
          break;
        case 'sheet':
          ensureSheet(op.name);
          break;
        case 'columns':
          applyColumns(ensureSheet(op.sheet), op.columns);
          break;
        case 'row':
          setRowValues(ensureSheet(op.sheet), op.values);
          break;
        case 'rows':
          appendRowsDense(ensureSheet(op.sheet), op.values);
          break;
        default:
          break;
      }
    }
  } else {
    for (const op of optimized) {
      switch (op.op) {
        case 'meta':
          applyMeta(wb, op);
          break;
        case 'sheet':
          ensureSheet(op.name);
          break;
        case 'columns':
          applyColumns(ensureSheet(op.sheet), op.columns);
          break;
        case 'row':
          setRowValues(ensureSheet(op.sheet), op.values);
          break;
        case 'rows':
          // Still batch when possible even on the general path (fused by optimizeOps).
          appendRowsDense(ensureSheet(op.sheet), op.values);
          break;
        case 'cell': {
          const ws = ensureSheet(op.sheet);
          if (op.style) {
            const cell = ws.getCell(op.address);
            cell.value = op.value;
            applyStyleToRange(ws, op.address, op.style);
          } else {
            // Compact path: avoid allocating full Cell for style-free values.
            const decoded = colCache.decodeAddress(op.address);
            if (decoded?.row && decoded?.col) {
              const row = ws.getRow(decoded.row);
              row._setCompact(decoded.col, op.value);
            } else {
              ws.getCell(op.address).value = op.value;
            }
          }
          break;
        }
        case 'cells': {
          const ws = ensureSheet(op.sheet);
          for (const [address, value] of Object.entries(op.map)) {
            const decoded = colCache.decodeAddress(address);
            if (decoded?.row && decoded?.col) {
              ws.getRow(decoded.row)._setCompact(decoded.col, value);
            } else {
              ws.getCell(address).value = value;
            }
          }
          break;
        }
        case 'style':
          applyStyleToRange(ensureSheet(op.sheet), op.range, op.style);
          break;
        case 'merge':
          ensureSheet(op.sheet).mergeCells(op.range);
          break;
        case 'views': {
          const ws = ensureSheet(op.sheet);
          ws.views = op.views as typeof ws.views;
          break;
        }
        case 'pageSetup': {
          const ws = ensureSheet(op.sheet);
          ws.pageSetup = Object.assign({}, ws.pageSetup, op.pageSetup);
          break;
        }
        case 'headerFooter': {
          const ws = ensureSheet(op.sheet);
          ws.headerFooter = Object.assign({}, ws.headerFooter, op.headerFooter);
          break;
        }
        case 'dataValidation': {
          const ws = ensureSheet(op.sheet);
          ws.dataValidations.add(op.address, op.rules);
          break;
        }
        case 'conditionalFormatting': {
          ensureSheet(op.sheet).addConditionalFormatting(op.cf);
          break;
        }
        case 'note': {
          ensureSheet(op.sheet).getCell(op.address).note = op.note;
          break;
        }
        case 'protect':
          applyProtect(ensureSheet(op.sheet), op.password, op.options);
          break;
        case 'sheetProtection':
          ensureSheet(op.sheet).sheetProtection = op.model;
          break;
        case 'table':
          ensureSheet(op.sheet).addTable(op.table);
          break;
        case 'media': {
          const realId = wb.addImage(op.image);
          mediaIdMap.set(op.id, realId);
          break;
        }
        case 'sheetImage': {
          const ws = ensureSheet(op.sheet);
          const realId = mediaIdMap.has(op.imageId) ? mediaIdMap.get(op.imageId)! : op.imageId;
          ws.addImage(realId, op.range);
          break;
        }
        case 'definedName':
          applyDefinedName(wb, op.name, op.refersTo);
          break;
        default:
          break;
      }
    }
  }

  // Ensure at least one worksheet (Excel expects it)
  if (sheets.size === 0) {
    wb.addWorksheet('Sheet1');
  }

  return wb;
}

/** @internal — for tests / diagnostics */
export function _isDenseMaterialize(ops: BuilderOp[]): boolean {
  return isDenseRectangularOps(optimizeOps(ops));
}
