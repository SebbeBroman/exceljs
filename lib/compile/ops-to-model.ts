/**
 * Compile builder ops into a plain Workbook snapshot.
 * Used by `build()`; write path materializes via the doc adapter instead.
 */

import type {
  CellValue,
  ColumnInput,
  ConditionalFormattingOptions,
  DataValidation,
  DefinedNameEntry,
  HeaderFooter,
  MediaImage,
  NoteValue,
  PageSetup,
  ProtectConfig,
  RowInput,
  SheetCell,
  SheetImagePlacement,
  SheetModel,
  SheetRow,
  Style,
  TableProperties,
  Workbook,
  WorkbookMeta,
  WorksheetViewInput,
} from '../model/types.js';
import type {BuilderOp} from '../builder/ops.js';
import colCache from '../utils/col-cache.js';

interface MutableSheet {
  id: number;
  name: string;
  columns?: ColumnInput[];
  merges: string[];
  /** row number → col → cell */
  grid: Map<number, Map<number, SheetCell>>;
  nextRow: number;
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  dataValidations?: Record<string, DataValidation>;
  conditionalFormattings?: ConditionalFormattingOptions[];
  notes?: Record<string, NoteValue>;
  protect?: ProtectConfig;
  sheetProtection?: Record<string, unknown>;
  tables?: TableProperties[];
  images?: SheetImagePlacement[];
}

function emptySheet(id: number, name: string): MutableSheet {
  return {id, name, merges: [], grid: new Map(), nextRow: 1};
}

function getOrCreateRow(sheet: MutableSheet, row: number): Map<number, SheetCell> {
  let r = sheet.grid.get(row);
  if (!r) {
    r = new Map();
    sheet.grid.set(row, r);
  }
  return r;
}

function setCell(sheet: MutableSheet, row: number, col: number, value: CellValue, style?: Style): void {
  const map = getOrCreateRow(sheet, row);
  const prev = map.get(col);
  map.set(col, {
    value,
    style: style ?? prev?.style,
  });
  if (row >= sheet.nextRow) sheet.nextRow = row + 1;
}

function applyStyleRange(sheet: MutableSheet, range: string, style: Style): void {
  const decoded = colCache.decode(range) as {
    top?: number;
    left?: number;
    bottom?: number;
    right?: number;
    row?: number;
    col?: number;
  };
  if (decoded.top != null && decoded.left != null) {
    for (let r = decoded.top; r <= (decoded.bottom as number); r++) {
      for (let c = decoded.left; c <= (decoded.right as number); c++) {
        const map = getOrCreateRow(sheet, r);
        const prev = map.get(c);
        map.set(c, {
          value: prev?.value ?? null,
          style: {...(prev?.style ?? {}), ...style},
        });
      }
    }
    return;
  }
  if (decoded.row != null && decoded.col != null) {
    const map = getOrCreateRow(sheet, decoded.row);
    const prev = map.get(decoded.col);
    map.set(decoded.col, {
      value: prev?.value ?? null,
      style: {...(prev?.style ?? {}), ...style},
    });
  }
}

function rowInputToPairs(
  values: RowInput,
  columns: ColumnInput[] | undefined,
): Array<[number, CellValue]> {
  if (Array.isArray(values)) {
    return values.map((v, i) => [i + 1, v as CellValue]);
  }
  const obj = values as Record<string, CellValue>;
  const pairs: Array<[number, CellValue]> = [];
  if (columns?.length) {
    columns.forEach((col, i) => {
      if (col.key && col.key in obj) {
        pairs.push([i + 1, obj[col.key]]);
      }
    });
  } else {
    // fallback: insertion order of object values into consecutive cols
    let i = 1;
    for (const v of Object.values(obj)) {
      pairs.push([i++, v]);
    }
  }
  return pairs;
}

function appendRow(sheet: MutableSheet, values: RowInput): void {
  const rowNum = sheet.nextRow;
  for (const [col, value] of rowInputToPairs(values, sheet.columns)) {
    setCell(sheet, rowNum, col, value);
  }
  // ensure nextRow advances even for empty row
  if (sheet.nextRow <= rowNum) sheet.nextRow = rowNum + 1;
}

function fuseMetaAndRows(ops: BuilderOp[]): {meta: WorkbookMeta; fused: BuilderOp[]; sawCellOps: boolean} {
  let meta: WorkbookMeta = {};
  const fused: BuilderOp[] = [];
  let sawCellOps = false;

  // Open row-fusion buffer: grow with push (O(n) total), flush on sheet change / other ops.
  let rowBufSheet: string | null = null;
  let rowBuf: RowInput[] | null = null;
  /** True when rowBuf is a shared reference to a single input `.rows()` values array. */
  let rowBufShared = false;

  const flushRows = (): void => {
    if (!rowBufSheet || !rowBuf || rowBuf.length === 0) {
      rowBufSheet = null;
      rowBuf = null;
      rowBufShared = false;
      return;
    }
    fused.push({op: 'rows', sheet: rowBufSheet, values: rowBuf});
    rowBufSheet = null;
    rowBuf = null;
    rowBufShared = false;
  };

  const appendRows = (sheet: string, values: readonly RowInput[]): void => {
    if (values.length === 0) return;
    if (rowBufSheet === sheet && rowBuf) {
      if (rowBufShared) {
        // Detach from the shared input array before mutating.
        rowBuf = rowBuf.slice();
        rowBufShared = false;
      }
      for (let i = 0; i < values.length; i++) rowBuf.push(values[i]!);
      return;
    }
    flushRows();
    // Single bulk `.rows()`: keep the caller's array (no copy) until a later append forces detach.
    rowBufSheet = sheet;
    rowBuf = values as RowInput[];
    rowBufShared = true;
  };

  const appendOneRow = (sheet: string, values: RowInput): void => {
    if (rowBufSheet === sheet && rowBuf) {
      if (rowBufShared) {
        rowBuf = rowBuf.slice();
        rowBufShared = false;
      }
      rowBuf.push(values);
      return;
    }
    flushRows();
    rowBufSheet = sheet;
    rowBuf = [values];
    rowBufShared = false;
  };

  for (const op of ops) {
    if (op.op === 'meta') {
      meta = {...meta, ...op.meta};
      continue;
    }
    if (op.op === 'rows') {
      appendRows(op.sheet, op.values);
      continue;
    }
    if (op.op === 'row') {
      appendOneRow(op.sheet, op.values);
      continue;
    }
    // Any non-row op ends the current fusion streak.
    flushRows();
    if (op.op === 'cells') {
      sawCellOps = true;
      if (Object.keys(op.map).length === 0) continue;
      fused.push(op);
      continue;
    }
    if (op.op === 'cell') {
      sawCellOps = true;
      fused.push(op);
      continue;
    }
    fused.push(op);
  }
  flushRows();

  return {meta, fused, sawCellOps};
}

/**
 * Op-log passes:
 * 1. Merge meta (last write wins per field)
 * 2. Drop empty ops (empty rows/cells)
 * 3. Fuse consecutive `row` / `rows` on the same sheet into one `rows`
 * 4. Last-write-wins for cell addresses (`cell` / `cells`) — skipped when none
 *
 * Never mutates the input op objects or their nested arrays/maps.
 */
export function optimizeOps(ops: BuilderOp[]): BuilderOp[] {
  const {meta, fused, sawCellOps} = fuseMetaAndRows(ops);

  let out: BuilderOp[];

  if (!sawCellOps) {
    // Dense / no random cells: skip LWW staging.
    out = fused;
  } else {
    // Last-write-wins for the same sheet + A1 address across cell / cells ops.
    const lastIndex = new Map<string, number>();
    const staged: Array<BuilderOp | null> = [];

    const cellKey = (sheet: string, address: string) => `${sheet}!${address}`;

    const clearPriorCell = (sheet: string, address: string): void => {
      const key = cellKey(sheet, address);
      const prevIdx = lastIndex.get(key);
      if (prevIdx === undefined) return;
      const prev = staged[prevIdx];
      if (!prev) return;
      if (prev.op === 'cell') {
        staged[prevIdx] = null;
      } else if (prev.op === 'cells') {
        const map = {...prev.map};
        delete map[address];
        staged[prevIdx] = Object.keys(map).length
          ? {op: 'cells', sheet: prev.sheet, map}
          : null;
      }
    };

    for (const op of fused) {
      if (op.op === 'cell') {
        clearPriorCell(op.sheet, op.address);
        lastIndex.set(cellKey(op.sheet, op.address), staged.length);
        staged.push(op);
        continue;
      }
      if (op.op === 'cells') {
        const map: Record<string, CellValue> = {};
        for (const [address, value] of Object.entries(op.map)) {
          clearPriorCell(op.sheet, address);
          map[address] = value;
          lastIndex.set(cellKey(op.sheet, address), staged.length);
        }
        if (Object.keys(map).length === 0) continue;
        staged.push({op: 'cells', sheet: op.sheet, map});
        continue;
      }
      staged.push(op);
    }

    out = staged.filter((op): op is BuilderOp => op != null);
  }

  if (Object.keys(meta).length) {
    out = [{op: 'meta', meta}, ...out];
  }
  return out;
}

/**
 * True when the op log is rectangular-append only: sheets/meta/columns and
 * sequential row(s) — no random cell patches, styles, or merges.
 * Dense materialize uses bulk `addRows` without per-cell random access.
 */
export function isDenseRectangularOps(ops: BuilderOp[]): boolean {
  for (const op of ops) {
    switch (op.op) {
      case 'meta':
      case 'sheet':
      case 'columns':
      case 'row':
      case 'rows':
        break;
      default:
        return false;
    }
  }
  return true;
}

export function compileToPlainWorkbook(ops: BuilderOp[]): Workbook {
  const optimized = optimizeOps(ops);
  let meta: WorkbookMeta = {};
  const sheets = new Map<string, MutableSheet>();
  let nextId = 1;
  const media: MediaImage[] = [];
  const mediaIdMap = new Map<number, number>();
  const definedNames: DefinedNameEntry[] = [];

  const ensure = (name: string): MutableSheet => {
    let s = sheets.get(name);
    if (!s) {
      s = emptySheet(nextId++, name);
      sheets.set(name, s);
    }
    return s;
  };

  for (const op of optimized) {
    switch (op.op) {
      case 'meta':
        meta = {...meta, ...op.meta};
        break;
      case 'sheet':
        ensure(op.name);
        break;
      case 'columns': {
        const s = ensure(op.sheet);
        s.columns = op.columns;
        // header row if any column has header
        const headers = op.columns.map(c =>
          Array.isArray(c.header) ? c.header[0] : (c.header as string | undefined),
        );
        if (headers.some(h => h != null && h !== '')) {
          appendRow(s, headers.map(h => h ?? null));
        }
        break;
      }
      case 'row':
        appendRow(ensure(op.sheet), op.values);
        break;
      case 'rows': {
        const s = ensure(op.sheet);
        for (const row of op.values) appendRow(s, row);
        break;
      }
      case 'cell': {
        const s = ensure(op.sheet);
        const addr = colCache.decodeAddress(op.address);
        setCell(s, addr.row as number, addr.col as number, op.value, op.style);
        break;
      }
      case 'cells': {
        const s = ensure(op.sheet);
        for (const [address, value] of Object.entries(op.map)) {
          const addr = colCache.decodeAddress(address);
          setCell(s, addr.row as number, addr.col as number, value);
        }
        break;
      }
      case 'style':
        applyStyleRange(ensure(op.sheet), op.range, op.style);
        break;
      case 'merge':
        ensure(op.sheet).merges.push(op.range);
        break;
      case 'views':
        ensure(op.sheet).views = op.views;
        break;
      case 'pageSetup': {
        const s = ensure(op.sheet);
        s.pageSetup = {...(s.pageSetup ?? {}), ...op.pageSetup};
        break;
      }
      case 'headerFooter': {
        const s = ensure(op.sheet);
        s.headerFooter = {...(s.headerFooter ?? {}), ...op.headerFooter};
        break;
      }
      case 'dataValidation': {
        const s = ensure(op.sheet);
        if (!s.dataValidations) s.dataValidations = {};
        s.dataValidations[op.address] = op.rules;
        break;
      }
      case 'conditionalFormatting': {
        const s = ensure(op.sheet);
        if (!s.conditionalFormattings) s.conditionalFormattings = [];
        s.conditionalFormattings.push(op.cf);
        break;
      }
      case 'note': {
        const s = ensure(op.sheet);
        if (!s.notes) s.notes = {};
        s.notes[op.address] = op.note;
        break;
      }
      case 'protect':
        ensure(op.sheet).protect = {password: op.password, options: op.options};
        break;
      case 'sheetProtection':
        ensure(op.sheet).sheetProtection = op.model;
        break;
      case 'table': {
        const s = ensure(op.sheet);
        if (!s.tables) s.tables = [];
        s.tables.push(op.table);
        break;
      }
      case 'media': {
        const idx = media.length;
        media.push(op.image);
        mediaIdMap.set(op.id, idx);
        break;
      }
      case 'sheetImage': {
        const s = ensure(op.sheet);
        if (!s.images) s.images = [];
        const imageId = mediaIdMap.has(op.imageId) ? mediaIdMap.get(op.imageId)! : op.imageId;
        s.images.push({imageId, range: op.range});
        break;
      }
      case 'definedName':
        definedNames.push({name: op.name, refersTo: op.refersTo});
        break;
      default:
        break;
    }
  }

  // preserve sheet creation order
  const sheetList: SheetModel[] = [];
  for (const s of sheets.values()) {
    const rows: SheetRow[] = [];
    const rowNums = [...s.grid.keys()].sort((a, b) => a - b);
    for (const number of rowNums) {
      const cellsMap = s.grid.get(number)!;
      const cells: Record<number, SheetCell> = {};
      for (const [col, cell] of cellsMap) {
        cells[col] = cell;
      }
      rows.push({number, cells});
    }
    const sheet: SheetModel = {
      id: s.id,
      name: s.name,
      rows,
      columns: s.columns,
      merges: s.merges.length ? s.merges : undefined,
    };
    if (s.views?.length) sheet.views = s.views;
    if (s.pageSetup && Object.keys(s.pageSetup).length) sheet.pageSetup = s.pageSetup;
    if (s.headerFooter && Object.keys(s.headerFooter).length) sheet.headerFooter = s.headerFooter;
    if (s.dataValidations && Object.keys(s.dataValidations).length) {
      sheet.dataValidations = s.dataValidations;
    }
    if (s.conditionalFormattings?.length) sheet.conditionalFormattings = s.conditionalFormattings;
    if (s.notes && Object.keys(s.notes).length) sheet.notes = s.notes;
    if (s.protect) sheet.protect = s.protect;
    if (s.sheetProtection) sheet.sheetProtection = s.sheetProtection;
    if (s.tables?.length) sheet.tables = s.tables;
    if (s.images?.length) sheet.images = s.images;
    sheetList.push(sheet);
  }

  const result: Workbook = {meta, sheets: sheetList};
  if (media.length) result.media = media;
  if (definedNames.length) result.definedNames = definedNames;
  return result;
}
