/**
 * Public streaming write API for Node (`streamWrite` from `@sebbebroman/exceljs/node`).
 * Thin ergonomic wrapper around the internal WorkbookWriter (not re-exported).
 */

import type {Writable} from 'node:stream';
import {once} from 'node:events';
import type WorksheetWriter from './worksheet-writer.js';
import type {
  ColumnInput,
  HeaderFooter,
  PageSetup,
  RowInput,
  WorksheetViewInput,
} from '../../model/types.js';
import WorkbookWriter, {
  type WorkbookWriterOptions,
  type ZipWriterOptions,
} from './workbook-writer.js';

/** Zip / shared-string / style options shared with the stream writer. */
export interface StreamWriteOptions {
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: Partial<ZipWriterOptions>;
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
}

/** Per-sheet options when opening a stream sheet. */
export interface StreamSheetOptions {
  columns?: ColumnInput[];
  state?: string;
  views?: WorksheetViewInput[];
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  autoFilter?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  properties?: any;
}

/**
 * Declarative sheet: name + row source (array, iterable, or async iterable).
 * Rows are committed as they are written so memory stays bounded.
 */
export interface StreamWriteSheetSpec extends StreamSheetOptions {
  name: string;
  rows: AsyncIterable<RowInput> | Iterable<RowInput>;
}

/** Declarative multi-sheet stream write. */
export interface StreamWriteDeclarative extends StreamWriteOptions {
  sheets: StreamWriteSheetSpec[];
}

/** Callback-style handle for one sheet. */
export interface StreamSheetHandle {
  /** Set column defs (headers/keys/widths). Prefer before rows. */
  columns(cols: ColumnInput[]): StreamSheetHandle;
  /** Append one row synchronously. Use rows() for backpressure with large sources. */
  row(values: RowInput): StreamSheetHandle;
  /** Append rows from an iterable, waiting for destination backpressure. */
  rows(values: AsyncIterable<RowInput> | Iterable<RowInput>): Promise<void>;
}

/** Callback-style workbook handle. */
export interface StreamWorkbookHandle {
  sheet(name: string, options?: StreamSheetOptions): StreamSheetHandle;
}

export type StreamWriteCallback = (w: StreamWorkbookHandle) => void | Promise<void>;

/**
 * Stream write specification: declarative object or callback for control flow.
 *
 * @example Declarative
 * ```ts
 * await streamWrite('out.xlsx', {
 *   useSharedStrings: true,
 *   sheets: [{ name: 'Data', columns: [...], rows: asyncRowSource() }],
 * });
 * ```
 *
 * @example Callback
 * ```ts
 * await streamWrite('out.xlsx', async w => {
 *   const sheet = w.sheet('Data', { columns: [...] });
 *   await sheet.rows(source);
 * });
 * ```
 */
export type StreamWriteSpec = StreamWriteDeclarative | StreamWriteCallback;

function hasAsyncIterator(
  rows: AsyncIterable<RowInput> | Iterable<RowInput>,
): rows is AsyncIterable<RowInput> {
  return (
    rows != null && typeof (rows as AsyncIterable<RowInput>)[Symbol.asyncIterator] === 'function'
  );
}

async function consumeRows(
  ws: WorksheetWriter,
  rows: AsyncIterable<RowInput> | Iterable<RowInput>,
): Promise<void> {
  if (hasAsyncIterator(rows)) {
    for await (const values of rows) {
      ws.writeRow(values);
      if (ws.stream.writableNeedDrain) await once(ws.stream, 'drain');
    }
    return;
  }
  for (const values of rows) {
    ws.writeRow(values);
    if (ws.stream.writableNeedDrain) await once(ws.stream, 'drain');
  }
}

/** Apply keys/widths without Column.header overwrite; append header row if needed. */
function applyColumns(ws: WorksheetWriter, columns: ColumnInput[]): void {
  const headers = columns.map(c => {
    if (c.header == null) return undefined;
    return Array.isArray(c.header) ? c.header[0] : c.header;
  });
  ws.columns = columns.map(c => ({
    key: c.key,
    width: c.width,
    hidden: c.hidden,
    style: c.style,
    outlineLevel: c.outlineLevel,
  }));
  if (headers.some(h => h != null && h !== '')) {
    ws.writeRow(headers.map(h => h ?? null));
  }
}

function applySheetOptions(ws: WorksheetWriter, options?: StreamSheetOptions): void {
  if (!options) return;
  if (options.columns) {
    applyColumns(ws, options.columns);
  }
  // state / views / pageSetup / headerFooter / autoFilter / properties
  // are passed into addWorksheet where supported; columns applied after.
}

function addWorksheetFromOptions(
  wb: WorkbookWriter,
  name: string,
  options?: StreamSheetOptions,
): WorksheetWriter {
  const opts = options ?? {};
  const ws = wb.addWorksheet(name, {
    state: opts.state,
    views: opts.views,
    pageSetup: opts.pageSetup,
    headerFooter: opts.headerFooter,
    autoFilter: opts.autoFilter,
    properties: opts.properties,
  });
  applySheetOptions(ws, opts);
  return ws;
}

class StreamSheetHandleImpl implements StreamSheetHandle {
  constructor(private readonly ws: WorksheetWriter) {}

  columns(cols: ColumnInput[]): StreamSheetHandle {
    applyColumns(this.ws, cols);
    return this;
  }

  row(values: RowInput): StreamSheetHandle {
    this.ws.writeRow(values);
    return this;
  }

  async rows(values: AsyncIterable<RowInput> | Iterable<RowInput>): Promise<void> {
    await consumeRows(this.ws, values);
  }
}

class StreamWorkbookHandleImpl implements StreamWorkbookHandle {
  constructor(private readonly wb: WorkbookWriter) {}

  sheet(name: string, options?: StreamSheetOptions): StreamSheetHandle {
    return new StreamSheetHandleImpl(addWorksheetFromOptions(this.wb, name, options));
  }
}

function writerOptionsFrom(
  dest: string | Writable,
  base: StreamWriteOptions | undefined,
): Partial<WorkbookWriterOptions> {
  const opts: Partial<WorkbookWriterOptions> = {...base};
  if (typeof dest === 'string') {
    opts.filename = dest;
  } else {
    opts.stream = dest;
  }
  return opts;
}

/**
 * Stream an xlsx workbook to a filesystem path or Node `Writable`.
 *
 * Prefer this over `writeFile` when row counts are large or data is produced
 * incrementally — rows are flushed as they are written.
 *
 * @param dest - File path or Writable stream
 * @param spec - Declarative `{ sheets, …options }` or async callback
 * @param options - When `spec` is a callback, write options (useStyles, etc.)
 */
export async function streamWrite(
  dest: string | Writable,
  spec: StreamWriteSpec,
  options?: StreamWriteOptions,
): Promise<void> {
  if (typeof spec === 'function') {
    const wb = new WorkbookWriter(writerOptionsFrom(dest, options));
    try {
      await spec(new StreamWorkbookHandleImpl(wb));
      await wb.commit();
    } catch (error) {
      wb.abort(error as Error);
      throw error;
    }
    return;
  }

  const {sheets, ...writeOpts} = spec;
  if (!sheets?.length) {
    throw new Error('streamWrite: declarative spec requires at least one sheet');
  }

  const wb = new WorkbookWriter(writerOptionsFrom(dest, writeOpts));
  try {
    for (const sheet of sheets) {
      const {name, rows, ...sheetOpts} = sheet;
      const ws = addWorksheetFromOptions(wb, name, sheetOpts);
      await consumeRows(ws, rows);
      ws.commit();
    }
    await wb.commit();
  } catch (error) {
    wb.abort(error as Error);
    throw error;
  }
}
