import type {
  CellValue,
  ColumnInput,
  ConditionalFormattingOptions,
  DataValidation,
  HeaderFooter,
  MediaImage,
  NoteValue,
  PageSetup,
  ProtectOptions,
  RowInput,
  SheetImageRange,
  SheetInit,
  SheetTitleInput,
  Style,
  TableProperties,
  Workbook,
  WorkbookInit,
  WorksheetViewInput,
  WriteOptions,
} from '../model/types.js';
import {type BuilderOp, emptyUsedFlags, noteFormulaValue, type UsedFlags} from './ops.js';
import colCache from '../utils/col-cache.js';
import {compileToPlainWorkbook} from '../compile/ops-to-model.js';
import {writeBuffer as encodeWriteBuffer} from '../xlsx/write-buffer.js';
import type {CsvStringifyOptions} from '../csv/public.js';

/** Duck-type BookView without importing the read graph (tree-shake). */
const WORKBOOK_VIEW_BRAND = Symbol.for('@sebbebroman/exceljs.WorkbookView');

function isWorkbookViewLike(value: unknown): value is {toJSON(): Workbook} {
  return Boolean(
    value &&
    typeof value === 'object' &&
    (value as {[k: symbol]: unknown})[WORKBOOK_VIEW_BRAND] === true &&
    typeof (value as {toJSON?: unknown}).toJSON === 'function',
  );
}

function normalizeTitle(title: SheetTitleInput): {text: string; style?: Style; merge?: string} {
  if (typeof title === 'string') return {text: title};
  return title;
}

function assertValidMerge(range: string): void {
  // Throws on non-A1 ranges so title({merge}) fails fast instead of
  // emitting a corrupt mergeCells entry at encode time.
  try {
    const decoded = colCache.decode(range) as {
      top?: number;
      left?: number;
      row?: number;
      col?: number;
    };
    const isRange = decoded.top != null && decoded.left != null;
    const isCell = decoded.row != null && decoded.col != null;
    if (!isRange && !isCell) throw new Error('not A1');
  } catch {
    throw new Error(`Invalid title merge range: ${range}`);
  }
}

export interface SheetBuilder {
  title(title: SheetTitleInput): SheetBuilder;
  row(values: RowInput): SheetBuilder;
  rows(values: RowInput[]): SheetBuilder;
  cell(address: string, value: CellValue, style?: Style): SheetBuilder;
  cells(map: Record<string, CellValue>): SheetBuilder;
  style(range: string, style: Style): SheetBuilder;
  merge(range: string): SheetBuilder;
  columns(cols: ColumnInput[]): SheetBuilder;
  views(views: WorksheetViewInput[]): SheetBuilder;
  pageSetup(setup: Partial<PageSetup>): SheetBuilder;
  headerFooter(hf: Partial<HeaderFooter>): SheetBuilder;
  dataValidation(address: string, rules: DataValidation): SheetBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): SheetBuilder;
  note(address: string, note: NoteValue): SheetBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): SheetBuilder;
  table(table: TableProperties): SheetBuilder;
  /** Place a workbook media image on this sheet. */
  image(imageId: number, range: SheetImageRange): SheetBuilder;
}

export interface WorkbookBuilder {
  sheet(name: string, init?: SheetInit | ((s: SheetBuilder) => void)): WorkbookBuilder;
  /**
   * Emit a title row on the active sheet (requires `.sheet(name)` first).
   * Prefer `SheetBuilder.title()` or `SheetInit.title` when working with
   * multiple sheets to avoid cursor mistakes.
   */
  title(title: SheetTitleInput): WorkbookBuilder;
  row(values: RowInput): WorkbookBuilder;
  rows(values: RowInput[]): WorkbookBuilder;
  cell(address: string, value: CellValue, style?: Style): WorkbookBuilder;
  cells(map: Record<string, CellValue>): WorkbookBuilder;
  style(range: string, style: Style): WorkbookBuilder;
  merge(range: string): WorkbookBuilder;
  columns(cols: ColumnInput[]): WorkbookBuilder;
  props(meta: WorkbookInit): WorkbookBuilder;

  views(views: WorksheetViewInput[]): WorkbookBuilder;
  pageSetup(setup: Partial<PageSetup>): WorkbookBuilder;
  headerFooter(hf: Partial<HeaderFooter>): WorkbookBuilder;
  dataValidation(address: string, rules: DataValidation): WorkbookBuilder;
  conditionalFormatting(cf: ConditionalFormattingOptions): WorkbookBuilder;
  note(address: string, note: NoteValue): WorkbookBuilder;
  /** Deferred: password hashed at materialize/encode time (chain stays sync). */
  protect(password?: string, options?: ProtectOptions): WorkbookBuilder;
  table(table: TableProperties): WorkbookBuilder;
  /**
   * Register workbook media. Returns the image id for `.image(id, range)` on a sheet.
   * Breaks the fluent chain by design (matches classic `addImage` → id pattern).
   */
  image(def: MediaImage): number;
  /** Place a previously registered image on the active sheet. */
  image(imageId: number, range: SheetImageRange): WorkbookBuilder;
  /** Alias for `image(def)` — clearer name for the register step. */
  addImage(def: MediaImage): number;
  /** Workbook-level defined name (`refersTo` e.g. `Sheet1!$A$1:$B$2`). */
  definedName(name: string, refersTo: string): WorkbookBuilder;

  /** Plain data snapshot (no I/O). */
  build(): Workbook;
  writeBuffer(opts?: WriteOptions): Promise<Uint8Array>;
  /**
   * Stringify the active (or first) sheet as CSV.
   * Implementation is loaded on demand so write-only bundles can tree-shake CSV.
   */
  csv(opts?: CsvStringifyOptions): Promise<string>;

  /** @internal */
  readonly _ops: BuilderOp[];
  /** @internal */
  readonly _used: UsedFlags;
}

class SheetBuilderImpl implements SheetBuilder {
  constructor(
    private readonly wb: WorkbookBuilderImpl,
    private readonly name: string,
  ) {}

  title(title: SheetTitleInput): SheetBuilder {
    this.wb.applyTitle(this.name, title);
    return this;
  }

  row(values: RowInput): SheetBuilder {
    this.wb._push({op: 'row', sheet: this.name, values});
    return this;
  }

  rows(values: RowInput[]): SheetBuilder {
    this.wb._push({op: 'rows', sheet: this.name, values});
    return this;
  }

  cell(address: string, value: CellValue, style?: Style): SheetBuilder {
    this.wb._pushCell(this.name, address, value, style);
    return this;
  }

  cells(map: Record<string, CellValue>): SheetBuilder {
    this.wb._push({op: 'cells', sheet: this.name, map});
    for (const v of Object.values(map)) noteFormulaValue(v, this.wb._used);
    return this;
  }

  style(range: string, style: Style): SheetBuilder {
    this.wb._used.styles = true;
    this.wb._push({op: 'style', sheet: this.name, range, style});
    return this;
  }

  merge(range: string): SheetBuilder {
    this.wb._used.merges = true;
    this.wb._push({op: 'merge', sheet: this.name, range});
    return this;
  }

  columns(cols: ColumnInput[]): SheetBuilder {
    this.wb._used.columns = true;
    if (cols.some(c => c.style)) this.wb._used.styles = true;
    this.wb._push({op: 'columns', sheet: this.name, columns: cols});
    return this;
  }

  views(views: WorksheetViewInput[]): SheetBuilder {
    this.wb._used.views = true;
    this.wb._push({op: 'views', sheet: this.name, views});
    return this;
  }

  pageSetup(setup: Partial<PageSetup>): SheetBuilder {
    this.wb._used.pageSetup = true;
    this.wb._push({op: 'pageSetup', sheet: this.name, pageSetup: setup});
    return this;
  }

  headerFooter(hf: Partial<HeaderFooter>): SheetBuilder {
    this.wb._used.pageSetup = true;
    this.wb._push({op: 'headerFooter', sheet: this.name, headerFooter: hf});
    return this;
  }

  dataValidation(address: string, rules: DataValidation): SheetBuilder {
    this.wb._used.dataValidations = true;
    this.wb._push({op: 'dataValidation', sheet: this.name, address, rules});
    return this;
  }

  conditionalFormatting(cf: ConditionalFormattingOptions): SheetBuilder {
    this.wb._used.conditionalFormatting = true;
    if (cf.rules?.some(r => r.style)) this.wb._used.styles = true;
    this.wb._push({op: 'conditionalFormatting', sheet: this.name, cf});
    return this;
  }

  note(address: string, note: NoteValue): SheetBuilder {
    this.wb._used.notes = true;
    this.wb._push({op: 'note', sheet: this.name, address, note});
    return this;
  }

  protect(password?: string, options?: ProtectOptions): SheetBuilder {
    this.wb._used.protection = true;
    this.wb._push({op: 'protect', sheet: this.name, password, options});
    return this;
  }

  table(table: TableProperties): SheetBuilder {
    this.wb._used.tables = true;
    this.wb._push({op: 'table', sheet: this.name, table});
    return this;
  }

  image(imageId: number, range: SheetImageRange): SheetBuilder {
    this.wb._used.images = true;
    this.wb._push({op: 'sheetImage', sheet: this.name, imageId, range});
    return this;
  }
}

class WorkbookBuilderImpl implements WorkbookBuilder {
  readonly _ops: BuilderOp[] = [];
  readonly _used: UsedFlags = emptyUsedFlags();
  private _cursor: string | null = null;
  private _nextMediaId = 0;

  constructor(init?: WorkbookInit | Workbook | {toJSON(): Workbook}) {
    if (!init) return;

    // viewWorkbook(...) handle → plain snapshot (no static import of read modules)
    if (isWorkbookViewLike(init)) {
      init = init.toJSON();
    }

    if (isPlainWorkbook(init)) {
      // Re-open plain Workbook: replay as ops for edit loop
      this._ops.push({op: 'meta', meta: init.meta});
      if (init.media?.length) {
        for (let i = 0; i < init.media.length; i++) {
          this._used.images = true;
          this._ops.push({op: 'media', id: i, image: init.media[i]!});
        }
        this._nextMediaId = init.media.length;
      }
      if (init.definedNames?.length) {
        for (const dn of init.definedNames) {
          this._used.definedNames = true;
          this._ops.push({op: 'definedName', name: dn.name, refersTo: dn.refersTo});
        }
      }
      for (const sheet of init.sheets) {
        this._ops.push({op: 'sheet', name: sheet.name});
        this._cursor = sheet.name;
        if (sheet.columns?.length) {
          this._used.columns = true;
          this._ops.push({op: 'columns', sheet: sheet.name, columns: sheet.columns});
        }
        for (const row of sheet.rows) {
          for (const [colStr, cell] of Object.entries(row.cells)) {
            const col = Number(colStr);
            const address = encodeCell(row.number, col);
            this._pushCell(sheet.name, address, cell.value, cell.style);
          }
        }
        for (const range of sheet.merges ?? []) {
          this._used.merges = true;
          this._ops.push({op: 'merge', sheet: sheet.name, range});
        }
        replaySheetFeatures(this, sheet);
      }
      return;
    }

    const {sheets, media, definedNames, ...rest} = init;
    if (Object.keys(rest).length) {
      this._ops.push({op: 'meta', meta: rest});
    }
    if (media?.length) {
      for (let i = 0; i < media.length; i++) {
        this._used.images = true;
        this._ops.push({op: 'media', id: i, image: media[i]!});
      }
      this._nextMediaId = media.length;
    }
    if (definedNames?.length) {
      for (const dn of definedNames) {
        this._used.definedNames = true;
        this._ops.push({op: 'definedName', name: dn.name, refersTo: dn.refersTo});
      }
    }
    if (sheets) {
      for (const sheet of sheets) {
        this.sheet(sheet.name);
        if (sheet.columns?.length) this.columns(sheet.columns);
        for (const row of sheet.rows) {
          for (const [colStr, cell] of Object.entries(row.cells)) {
            const address = encodeCell(row.number, Number(colStr));
            this.cell(address, cell.value, cell.style);
          }
        }
        for (const m of sheet.merges ?? []) this.merge(m);
        replaySheetFeatures(this, sheet);
      }
    }
  }

  _push(op: BuilderOp): void {
    this._ops.push(op);
  }

  _pushCell(sheet: string, address: string, value: CellValue, style?: Style): void {
    noteFormulaValue(value, this._used);
    if (style) this._used.styles = true;
    this._ops.push(
      style ? {op: 'cell', sheet, address, value, style} : {op: 'cell', sheet, address, value},
    );
  }

  private requireCursor(): string {
    if (!this._cursor) {
      throw new Error(
        'No active sheet. Call .sheet(name) before title/row/cell/style/merge/columns operations.',
      );
    }
    return this._cursor;
  }

  /** @internal Emit title row (+ optional merge/style) on a named sheet.
   * Style applies to the merge range when given, otherwise to A1 (the title cell).
   */
  applyTitle(sheet: string, title: SheetTitleInput): void {
    const {text, style, merge} = normalizeTitle(title);
    if (merge !== undefined) assertValidMerge(merge);
    this._ops.push({op: 'row', sheet, values: [text]});
    if (merge) {
      this._used.merges = true;
      this._ops.push({op: 'merge', sheet, range: merge});
    }
    if (style) {
      this._used.styles = true;
      this._ops.push({op: 'style', sheet, range: merge ?? 'A1', style});
    }
  }

  sheet(name: string, init?: SheetInit | ((s: SheetBuilder) => void)): WorkbookBuilder {
    this._ops.push({op: 'sheet', name});
    this._cursor = name;

    if (typeof init === 'function') {
      init(new SheetBuilderImpl(this, name));
    } else if (init) {
      if (init.title != null) this.applyTitle(name, init.title);
      if (init.columns?.length) this.columns(init.columns);
      if (init.rows?.length) this.rows(init.rows);
      if (init.merges?.length) {
        for (const m of init.merges) this.merge(m);
      }
      if (init.views?.length) this.views(init.views);
      if (init.pageSetup) this.pageSetup(init.pageSetup);
      if (init.headerFooter) this.headerFooter(init.headerFooter);
    }
    return this;
  }

  title(title: SheetTitleInput): WorkbookBuilder {
    this.applyTitle(this.requireCursor(), title);
    return this;
  }

  row(values: RowInput): WorkbookBuilder {
    this._ops.push({op: 'row', sheet: this.requireCursor(), values});
    return this;
  }

  rows(values: RowInput[]): WorkbookBuilder {
    this._ops.push({op: 'rows', sheet: this.requireCursor(), values});
    return this;
  }

  cell(address: string, value: CellValue, style?: Style): WorkbookBuilder {
    this._pushCell(this.requireCursor(), address, value, style);
    return this;
  }

  cells(map: Record<string, CellValue>): WorkbookBuilder {
    const sheet = this.requireCursor();
    this._ops.push({op: 'cells', sheet, map});
    for (const v of Object.values(map)) noteFormulaValue(v, this._used);
    return this;
  }

  style(range: string, style: Style): WorkbookBuilder {
    this._used.styles = true;
    this._ops.push({op: 'style', sheet: this.requireCursor(), range, style});
    return this;
  }

  merge(range: string): WorkbookBuilder {
    this._used.merges = true;
    this._ops.push({op: 'merge', sheet: this.requireCursor(), range});
    return this;
  }

  columns(cols: ColumnInput[]): WorkbookBuilder {
    this._used.columns = true;
    if (cols.some(c => c.style)) this._used.styles = true;
    this._ops.push({op: 'columns', sheet: this.requireCursor(), columns: cols});
    return this;
  }

  props(meta: WorkbookInit): WorkbookBuilder {
    const {sheets: _s, media: _m, definedNames: _d, ...rest} = meta;
    this._ops.push({op: 'meta', meta: rest});
    return this;
  }

  views(views: WorksheetViewInput[]): WorkbookBuilder {
    this._used.views = true;
    this._ops.push({op: 'views', sheet: this.requireCursor(), views});
    return this;
  }

  pageSetup(setup: Partial<PageSetup>): WorkbookBuilder {
    this._used.pageSetup = true;
    this._ops.push({op: 'pageSetup', sheet: this.requireCursor(), pageSetup: setup});
    return this;
  }

  headerFooter(hf: Partial<HeaderFooter>): WorkbookBuilder {
    this._used.pageSetup = true;
    this._ops.push({op: 'headerFooter', sheet: this.requireCursor(), headerFooter: hf});
    return this;
  }

  dataValidation(address: string, rules: DataValidation): WorkbookBuilder {
    this._used.dataValidations = true;
    this._ops.push({op: 'dataValidation', sheet: this.requireCursor(), address, rules});
    return this;
  }

  conditionalFormatting(cf: ConditionalFormattingOptions): WorkbookBuilder {
    this._used.conditionalFormatting = true;
    if (cf.rules?.some(r => r.style)) this._used.styles = true;
    this._ops.push({op: 'conditionalFormatting', sheet: this.requireCursor(), cf});
    return this;
  }

  note(address: string, note: NoteValue): WorkbookBuilder {
    this._used.notes = true;
    this._ops.push({op: 'note', sheet: this.requireCursor(), address, note});
    return this;
  }

  protect(password?: string, options?: ProtectOptions): WorkbookBuilder {
    this._used.protection = true;
    this._ops.push({op: 'protect', sheet: this.requireCursor(), password, options});
    return this;
  }

  table(table: TableProperties): WorkbookBuilder {
    this._used.tables = true;
    this._ops.push({op: 'table', sheet: this.requireCursor(), table});
    return this;
  }

  image(def: MediaImage): number;
  image(imageId: number, range: SheetImageRange): WorkbookBuilder;
  image(defOrId: MediaImage | number, range?: SheetImageRange): number | WorkbookBuilder {
    if (typeof defOrId === 'number') {
      if (range === undefined) {
        throw new Error('image(id, range): range is required when placing a sheet image');
      }
      this._used.images = true;
      this._ops.push({op: 'sheetImage', sheet: this.requireCursor(), imageId: defOrId, range});
      return this as WorkbookBuilder;
    }
    const id = this._nextMediaId++;
    this._used.images = true;
    this._ops.push({op: 'media', id, image: defOrId});
    return id;
  }

  addImage(def: MediaImage): number {
    return this.image(def);
  }

  definedName(name: string, refersTo: string): WorkbookBuilder {
    this._used.definedNames = true;
    this._ops.push({op: 'definedName', name, refersTo});
    return this;
  }

  build(): Workbook {
    return compileToPlainWorkbook(this._ops);
  }

  writeBuffer(opts?: WriteOptions): Promise<Uint8Array> {
    return encodeWriteBuffer(this, opts);
  }

  async csv(opts?: CsvStringifyOptions): Promise<string> {
    // Dynamic import: keep CSV out of write-only static graphs.
    const {stringifyCsv} = await import('../csv/public.js');
    const options: CsvStringifyOptions = opts ? {...opts} : {};
    if (options.sheetName == null && options.sheetId == null && this._cursor) {
      options.sheetName = this._cursor;
    }
    return stringifyCsv(this, options);
  }
}

function replaySheetFeatures(wb: WorkbookBuilderImpl, sheet: Workbook['sheets'][number]): void {
  const name = sheet.name;
  if (sheet.views?.length) {
    wb._used.views = true;
    wb._push({op: 'views', sheet: name, views: sheet.views});
  }
  if (sheet.pageSetup && Object.keys(sheet.pageSetup).length) {
    wb._used.pageSetup = true;
    wb._push({op: 'pageSetup', sheet: name, pageSetup: sheet.pageSetup});
  }
  if (sheet.headerFooter && Object.keys(sheet.headerFooter).length) {
    wb._used.pageSetup = true;
    wb._push({op: 'headerFooter', sheet: name, headerFooter: sheet.headerFooter});
  }
  if (sheet.dataValidations) {
    for (const [address, rules] of Object.entries(sheet.dataValidations)) {
      if (!rules) continue;
      wb._used.dataValidations = true;
      wb._push({op: 'dataValidation', sheet: name, address, rules});
    }
  }
  if (sheet.conditionalFormattings?.length) {
    for (const cf of sheet.conditionalFormattings) {
      wb._used.conditionalFormatting = true;
      if (cf.rules?.some(r => r.style)) wb._used.styles = true;
      wb._push({op: 'conditionalFormatting', sheet: name, cf});
    }
  }
  if (sheet.notes) {
    for (const [address, note] of Object.entries(sheet.notes)) {
      wb._used.notes = true;
      wb._push({op: 'note', sheet: name, address, note});
    }
  }
  if (sheet.protect) {
    wb._used.protection = true;
    wb._push({
      op: 'protect',
      sheet: name,
      password: sheet.protect.password,
      options: sheet.protect.options,
    });
  } else if (sheet.sheetProtection) {
    wb._used.protection = true;
    wb._push({op: 'sheetProtection', sheet: name, model: sheet.sheetProtection});
  }
  if (sheet.tables?.length) {
    for (const table of sheet.tables) {
      wb._used.tables = true;
      wb._push({op: 'table', sheet: name, table});
    }
  }
  if (sheet.images?.length) {
    for (const img of sheet.images) {
      wb._used.images = true;
      wb._push({op: 'sheetImage', sheet: name, imageId: img.imageId, range: img.range});
    }
  }
}

function isPlainWorkbook(value: WorkbookInit | Workbook): value is Workbook {
  return (
    typeof value === 'object' &&
    value !== null &&
    'meta' in value &&
    'sheets' in value &&
    Array.isArray((value as Workbook).sheets) &&
    typeof (value as Workbook).meta === 'object'
  );
}

/** Create a workbook builder (primary public API). */
/**
 * Create a write builder. Accepts:
 * - nothing / meta init
 * - plain `Workbook` snapshot (`build()` / `load()` / `view.toJSON()`)
 * - {@link WorkbookView} from `viewWorkbook` (`workbook(view)`)
 */
export function workbook(init?: WorkbookInit | Workbook | {toJSON(): Workbook}): WorkbookBuilder {
  return new WorkbookBuilderImpl(init);
}

/** True if value looks like a builder (has op log). */
export function isWorkbookBuilder(value: unknown): value is WorkbookBuilder {
  return Boolean(
    value &&
    typeof value === 'object' &&
    '_ops' in value &&
    Array.isArray((value as WorkbookBuilder)._ops),
  );
}

// --- address helpers (minimal; col-cache used at compile time for ranges) ---

function encodeCell(row: number, col: number): string {
  return `${colToLetter(col)}${row}`;
}

function colToLetter(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || 'A';
}
