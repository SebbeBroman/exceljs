/** Compile directly into the existing OOXML encoder's model, without a document graph. */
import type {BuilderOp} from '../builder/ops.js';
import type {Workbook, ColumnInput, RowInput, Style, ProtectOptions} from '../model/types.js';
import {optimizeOps} from './ops-to-model.js';
import {valueToModel, mergeStyles} from '../model/cell-model.js';
import colCache from '../utils/col-cache.js';
import {columnsToModel} from '../model/row-model.js';
import Range from '../model/range.js';
import DefinedNames from '../model/defined-names.js';
import Note from '../model/note.js';
import Enums from '../model/enums.js';
import Encryptor from '../utils/encryptor.js';
import type {CellValueModel} from '../model/cell-model.js';
import type {RowModelData} from '../model/xlsx-model.js';
import type {WorksheetModelData} from '../model/xlsx-model.js';

export interface XlsxWorkbookModel {
  worksheets: WorksheetModelData[];
  properties: Record<string, unknown>;
  [key: string]: unknown;
}
interface ModelRow extends RowModelData {
  slots: Array<CellValueModel | undefined>;
}
interface SheetState {
  model: WorksheetModelData;
  rows: Map<number, ModelRow>;
  columns: ColumnInput[];
  nextRow: number;
  merges: Range[];
}

function newSheet(id: number, name: string): SheetState {
  return {
    columns: [],
    rows: new Map(),
    nextRow: 1,
    merges: [],
    model: {
      id,
      name,
      state: 'visible',
      dataValidations: {},
      properties: {defaultRowHeight: 15, dyDescent: 55, outlineLevelCol: 0, outlineLevelRow: 0},
      pageSetup: {
        margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
        orientation: 'portrait',
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: false,
        pageOrder: 'downThenOver',
        blackAndWhite: false,
        draft: false,
        cellComments: 'None',
        errors: 'displayed',
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        showRowColHeaders: false,
        showGridLines: false,
        horizontalCentered: false,
        verticalCentered: false,
      },
      headerFooter: {differentFirst: false, differentOddEven: false},
      rowBreaks: [],
      views: [],
      autoFilter: null,
      media: [],
      sheetProtection: null,
      tables: [],
      pivotTables: [],
      conditionalFormattings: [],
      merges: [],
    },
  };
}
function ensureRow(sheet: SheetState, number: number): ModelRow {
  let row = sheet.rows.get(number);
  if (!row) {
    row = {
      number,
      cells: [],
      min: 0,
      max: 0,
      style: {},
      hidden: false,
      outlineLevel: 0,
      collapsed: false,
      slots: [],
    };
    sheet.rows.set(number, row);
    sheet.nextRow = Math.max(sheet.nextRow, number + 1);
  }
  return row;
}
function ensureCell(sheet: SheetState, row: number, col: number): CellValueModel {
  const r = ensureRow(sheet, row);
  let cell = r.slots[col - 1];
  if (!cell) {
    cell = valueToModel(colCache.encodeAddress(row, col), null);
    cell.style = mergeStyles(r.style, sheet.columns[col - 1]?.style, {});
    r.slots[col - 1] = cell;
  }
  return cell;
}
function setCell(
  sheet: SheetState,
  row: number,
  col: number,
  value: unknown,
  keepStyle = false,
): void {
  const r = ensureRow(sheet, row);
  const previous = r.slots[col - 1];
  if (keepStyle && previous?.type === Enums.ValueType.Merge) {
    const master = colCache.decodeAddress(previous.master as string);
    setCell(sheet, master.row!, master.col!, value, true);
    return;
  }
  if (value === undefined && !keepStyle) {
    delete r.slots[col - 1];
    return;
  }
  const cell = valueToModel(colCache.encodeAddress(row, col), value);
  const columnStyle = sheet.columns[col - 1]?.style;
  if (keepStyle && previous?.style) cell.style = previous.style;
  else if (columnStyle || Object.keys(r.style ?? {}).length)
    cell.style = mergeStyles(r.style, columnStyle, {});
  if (keepStyle && previous?.comment) cell.comment = previous.comment;
  r.slots[col - 1] = cell;
}
function appendRow(sheet: SheetState, values: RowInput): void {
  const number = sheet.nextRow;
  ensureRow(sheet, number);
  if (Array.isArray(values)) {
    // Legacy array rows accept a sparse one-based array when index zero is absent.
    const offset = Object.hasOwn(values, '0') ? 1 : 0;
    values.forEach((value, index) => {
      if (value !== undefined) setCell(sheet, number, index + offset, value);
    });
  } else {
    sheet.columns.forEach((column, index) => {
      if (column.key && (values as Record<string, unknown>)[column.key] !== undefined)
        setCell(sheet, number, index + 1, (values as Record<string, unknown>)[column.key]);
    });
  }
}
function patchStyle(target: Style, patch: Style): void {
  for (const key of ['font', 'alignment', 'border', 'protection'] as const) {
    if (patch[key]) Object.assign(target, {[key]: {...target[key], ...patch[key]}});
  }
  if (patch.fill) target.fill = patch.fill;
  if (patch.numFmt) target.numFmt = patch.numFmt;
}
function styleRange(sheet: SheetState, range: string, style: Style): void {
  if (/^\d+$/.test(range)) {
    const row = ensureRow(sheet, Number(range));
    // Legacy row styling assigns font even when the patch omits it, then
    // propagates each assigned row property to all existing cells.
    const rowStyle = row.style! as Style;
    rowStyle.font = style.font ? {...rowStyle.font, ...style.font} : rowStyle.font;
    const patch = {font: rowStyle.font} as Style;
    for (const key of ['alignment', 'border'] as const) {
      if (style[key]) Object.assign(patch, {[key]: {...rowStyle[key], ...style[key]}});
    }
    if (style.fill) patch.fill = style.fill;
    if (style.numFmt) patch.numFmt = style.numFmt;
    Object.assign(rowStyle, patch);
    for (const cell of row.slots) if (cell) Object.assign((cell.style ??= {}), patch);
    return;
  }
  const dimensions = new Range(range);
  for (let row = dimensions.top; row <= dimensions.bottom; row++) {
    for (let col = dimensions.left; col <= dimensions.right; col++) {
      const cell = ensureCell(sheet, row, col);
      patchStyle((cell.style ??= {}) as Style, style);
    }
  }
}
function mergeRange(sheet: SheetState, range: string): void {
  const dimensions = new Range(range);
  if (sheet.merges.some(merge => merge.intersects(dimensions)))
    throw new Error('Cannot merge already merged cells');
  const master = ensureCell(sheet, dimensions.top, dimensions.left);
  const style = (master.style ??= {});
  for (let row = dimensions.top; row <= dimensions.bottom; row++) {
    for (let col = dimensions.left; col <= dimensions.right; col++) {
      if (row === dimensions.top && col === dimensions.left) continue;
      ensureRow(sheet, row).slots[col - 1] = {
        address: colCache.encodeAddress(row, col),
        type: Enums.ValueType.Merge,
        master: master.address,
        style,
      };
    }
  }
  sheet.merges.push(dimensions);
  sheet.model.merges!.push(dimensions.range);
}
function protection(password?: string, options?: ProtectOptions): Record<string, unknown> {
  const model: Record<string, unknown> = {sheet: true};
  const spinCount =
    options?.spinCount == null
      ? 100000
      : Number.isFinite(options.spinCount)
        ? Math.round(Math.max(0, options.spinCount))
        : 100000;
  if (password) {
    const salt = Encryptor.randomBytesBase64(16);
    Object.assign(model, {
      algorithmName: 'SHA-512',
      saltValue: salt,
      spinCount,
      hashValue: Encryptor.convertPasswordToHash(password, 'SHA512', salt, spinCount),
    });
  }
  Object.assign(model, options);
  if (!password) delete model.spinCount;
  return model;
}
function finishSheet(sheet: SheetState): WorksheetModelData {
  const model = sheet.model;
  model.cols = columnsToModel(sheet.columns);
  model.rows = [];
  model.dimensions = new Range();
  for (const row of [...sheet.rows.values()].sort((a, b) => a.number - b.number)) {
    const cells: CellValueModel[] = [];
    let min = 0;
    let max = 0;
    for (let index = 0; index < row.slots.length; index++) {
      const cell = row.slots[index];
      if (!cell) continue;
      if (!min) min = index + 1;
      max = index + 1;
      cells.push(cell);
    }
    if (!cells.length && !row.height) continue;
    const {slots: _slots, ...result} = row;
    result.cells = cells;
    result.min = min;
    result.max = max;
    model.dimensions.expand(row.number, result.min, row.number, result.max);
    model.rows.push(result);
  }
  return model;
}

export async function compileToXlsxModel(
  ops: BuilderOp[],
  snapshot?: Workbook,
): Promise<XlsxWorkbookModel> {
  const sheets = new Map<string, SheetState>();
  const now = new Date();
  const model: XlsxWorkbookModel = {
    worksheets: [],
    creator: 'Unknown',
    lastModifiedBy: 'Unknown',
    created: now,
    modified: now,
    properties: {},
    views: [],
    media: [],
    pivotTables: [],
    calcProperties: {},
    company: '',
    manager: '',
    title: '',
    subject: '',
    description: '',
    keywords: '',
    category: '',
  };
  const media: Array<Record<string, unknown>> = [];
  const mediaIds = new Map<number, number>();
  const names = new DefinedNames();
  const ensure = (name: string): SheetState => {
    let sheet = sheets.get(name);
    if (!sheet) {
      if (
        typeof name !== 'string' ||
        !name ||
        name === 'History' ||
        /[*?:/\\[\]]/.test(name) ||
        /(^')|('$)/.test(name)
      )
        throw new Error(`Invalid worksheet name: ${name}`);
      const actual = name.slice(0, 31);
      if ([...sheets.values()].some(s => s.model.name!.toLowerCase() === actual.toLowerCase()))
        throw new Error(`Worksheet name already exists: ${actual}`);
      sheet = newSheet(sheets.size + 1, actual);
      sheets.set(name, sheet);
      const source = snapshot?.sheets.find(s => s.name === name);
      if (source) {
        sheet.columns = structuredClone(source.columns ?? []);
        for (const input of source.rows) {
          const row = ensureRow(sheet, input.number);
          row.height = input.height;
          row.hidden = input.hidden;
          row.style = structuredClone(input.style ?? {});
          for (const [col, cell] of Object.entries(input.cells)) {
            setCell(sheet, input.number, Number(col), cell.value);
            if (cell.style)
              patchStyle(
                (ensureCell(sheet, input.number, Number(col)).style ??= {}) as Style,
                structuredClone(cell.style),
              );
          }
        }
      }
    }
    return sheet;
  };
  for (const op of optimizeOps(ops)) {
    switch (op.op) {
      case 'meta':
        for (const [key, value] of Object.entries(op.meta)) {
          if (value != null) model[key] = structuredClone(value);
        }
        break;
      case 'sheet':
        ensure(op.name);
        break;
      case 'columns': {
        const sheet = ensure(op.sheet);
        sheet.columns = structuredClone(op.columns);
        const headers = op.columns.map(c => (Array.isArray(c.header) ? c.header[0] : c.header));
        if (headers.some(h => h != null && h !== ''))
          appendRow(
            sheet,
            headers.map(h => h ?? null),
          );
        break;
      }
      case 'row':
        appendRow(ensure(op.sheet), op.values);
        break;
      case 'rows': {
        const sheet = ensure(op.sheet);
        for (const row of op.values) appendRow(sheet, row);
        break;
      }
      case 'cell': {
        const sheet = ensure(op.sheet);
        const address = colCache.decodeAddress(op.address);
        setCell(sheet, address.row!, address.col!, op.value, !!op.style);
        if (op.style) styleRange(sheet, op.address, op.style);
        break;
      }
      case 'cells': {
        const sheet = ensure(op.sheet);
        for (const [address, value] of Object.entries(op.map)) {
          const a = colCache.decodeAddress(address);
          setCell(sheet, a.row!, a.col!, value);
        }
        break;
      }
      case 'style':
        styleRange(ensure(op.sheet), op.range, op.style);
        break;
      case 'merge':
        mergeRange(ensure(op.sheet), op.range);
        break;
      case 'views':
        ensure(op.sheet).model.views = structuredClone(op.views) as never;
        break;
      case 'pageSetup':
        Object.assign(ensure(op.sheet).model.pageSetup!, structuredClone(op.pageSetup));
        break;
      case 'headerFooter':
        Object.assign(ensure(op.sheet).model.headerFooter!, structuredClone(op.headerFooter));
        break;
      case 'dataValidation':
        ensure(op.sheet).model.dataValidations![op.address] = structuredClone(op.rules);
        break;
      case 'conditionalFormatting':
        ensure(op.sheet).model.conditionalFormattings!.push(structuredClone(op.cf));
        break;
      case 'note': {
        const a = colCache.decodeAddress(op.address);
        ensureCell(ensure(op.sheet), a.row!, a.col!).comment = new Note(
          structuredClone(op.note),
        ).model;
        break;
      }
      case 'protect':
        ensure(op.sheet).model.sheetProtection = protection(op.password, op.options);
        break;
      case 'sheetProtection':
        ensure(op.sheet).model.sheetProtection = {...op.model};
        break;
      case 'table': {
        const sheet = ensure(op.sheet);
        const {placeTable} = await import('../model/table.js');
        const table = placeTable(structuredClone(op.table), (row, col, value, style) => {
          setCell(sheet, row, col, value, true);
          if (style) Object.assign((ensureCell(sheet, row, col).style ??= {}), style);
        });
        const existing = sheet.model.tables!.findIndex(t => t.name === table.name);
        if (existing < 0) sheet.model.tables!.push(table);
        else sheet.model.tables![existing] = table;
        break;
      }
      case 'media':
        mediaIds.set(op.id, media.length);
        media.push({...op.image, type: 'image'});
        break;
      case 'sheetImage': {
        const sheet = ensure(op.sheet);
        const {default: Image} = await import('../model/image.js');
        const image = new Image(
          {
            getColumn: col => ({
              width: sheet.columns[col - 1]?.width,
              isCustomWidth:
                sheet.columns[col - 1]?.width != null && sheet.columns[col - 1]?.width !== 9,
            }),
            getRow: row => sheet.rows.get(row),
          },
          {type: 'image', imageId: mediaIds.get(op.imageId) ?? op.imageId, range: op.range},
        );
        sheet.model.media!.push(image.model);
        break;
      }
      case 'definedName':
        try {
          names.add(op.refersTo, op.name);
        } catch {
          names.model = [
            ...names.model.filter(d => d.name !== op.name),
            {name: op.name, ranges: [op.refersTo]},
          ];
        }
        break;
    }
  }
  if (!sheets.size) ensure('Sheet1');
  model.creator ||= 'Unknown';
  model.lastModifiedBy ||= 'Unknown';
  model.worksheets = [...sheets.values()].map(finishSheet);
  model.media = media;
  model.definedNames = names.model;
  return model;
}

/** Normalize a plain snapshot without creating per-cell builder operations. */
export async function plainToXlsxModel(workbook: Workbook): Promise<XlsxWorkbookModel> {
  const ops: BuilderOp[] = [{op: 'meta', meta: workbook.meta}];
  workbook.media?.forEach((image, id) => ops.push({op: 'media', id, image}));
  workbook.definedNames?.forEach(d =>
    ops.push({op: 'definedName', name: d.name, refersTo: d.refersTo}),
  );
  for (const sheet of workbook.sheets) {
    ops.push({op: 'sheet', name: sheet.name});
    if (sheet.views) ops.push({op: 'views', sheet: sheet.name, views: sheet.views});
    if (sheet.pageSetup) ops.push({op: 'pageSetup', sheet: sheet.name, pageSetup: sheet.pageSetup});
    if (sheet.headerFooter)
      ops.push({op: 'headerFooter', sheet: sheet.name, headerFooter: sheet.headerFooter});
    for (const range of sheet.merges ?? []) ops.push({op: 'merge', sheet: sheet.name, range});
    for (const [address, rules] of Object.entries(sheet.dataValidations ?? {}))
      ops.push({op: 'dataValidation', sheet: sheet.name, address, rules});
    for (const cf of sheet.conditionalFormattings ?? [])
      ops.push({op: 'conditionalFormatting', sheet: sheet.name, cf});
    for (const [address, note] of Object.entries(sheet.notes ?? {}))
      ops.push({op: 'note', sheet: sheet.name, address, note});
    if (sheet.protect) ops.push({op: 'protect', sheet: sheet.name, ...sheet.protect});
    else if (sheet.sheetProtection)
      ops.push({op: 'sheetProtection', sheet: sheet.name, model: sheet.sheetProtection});
    for (const table of sheet.tables ?? []) ops.push({op: 'table', sheet: sheet.name, table});
    for (const image of sheet.images ?? [])
      ops.push({op: 'sheetImage', sheet: sheet.name, ...image});
  }
  return compileToXlsxModel(ops, workbook);
}
