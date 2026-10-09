import {EventEmitter} from 'node:events';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {Readable} from 'node:stream';
import {eachSaxChunk} from '../../utils/parse-sax.js';
import streamZipEntries from '../../utils/stream-zip-reader.js';
import StyleManager from '../../xlsx/parser/style/styles-xform.js';
import WorkbookXform from '../../xlsx/parser/book/workbook-xform.js';
import RelationshipsXform from '../../xlsx/parser/core/relationships-xform.js';
import WorksheetReader from './worksheet-reader.js';
import HyperlinkReader from './hyperlink-reader.js';

export interface WorkbookStreamReaderOptions {
  worksheets?: 'emit' | 'ignore' | string;
  sharedStrings?: 'cache' | 'emit' | 'ignore' | string;
  hyperlinks?: 'cache' | 'emit' | 'ignore' | string;
  styles?: 'cache' | 'ignore' | string;
  entries?: 'emit' | 'ignore' | string;
}

export type WorkbookReaderInput =
  | string
  | Readable
  | AsyncIterable<unknown>
  | NodeJS.ReadableStream;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type FontState = Record<string, any> | null;

class WorkbookReader extends EventEmitter {
  input: WorkbookReaderInput;
  options: WorkbookStreamReaderOptions;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  styles: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  stream?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbookRels?: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  properties?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  model?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sharedStrings?: any[];

  static Options: {
    worksheets: string[];
    sharedStrings: string[];
    hyperlinks: string[];
    styles: string[];
    entries: string[];
  };

  constructor(input: WorkbookReaderInput, options: Partial<WorkbookStreamReaderOptions> = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _getStream(input: WorkbookReaderInput): any {
    if (
      input instanceof Readable ||
      (input && typeof (input as AsyncIterable<unknown>)[Symbol.asyncIterator] === 'function')
    ) {
      return input;
    }
    // Duck-type legacy event-emitter streams (data/end/error)
    if (input && typeof (input as NodeJS.ReadableStream).on === 'function') {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  async read(
    input?: WorkbookReaderInput,
    options?: Partial<WorkbookStreamReaderOptions>,
  ): Promise<void> {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator](): AsyncGenerator<InstanceType<typeof WorksheetReader>> {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async *parse(
    input?: WorkbookReaderInput,
    options?: Partial<WorkbookStreamReaderOptions>,
  ): AsyncGenerator<{eventType: string; value: any}> {
    if (options) this.options = options as WorkbookStreamReaderOptions;
    const stream = (this.stream = this._getStream(input || this.input));

    // worksheets deferred until shared strings + workbook rels are ready
    const waitingWorkSheets: {sheetNo: string; path: string}[] = [];
    let tempDir: string | null = null;

    try {
      for await (const entry of streamZipEntries(stream)) {
        let match: RegExpMatchArray | null;
        let sheetNo: string;
        switch (entry.path) {
          case '_rels/.rels':
            break;
          case 'xl/_rels/workbook.xml.rels':
            await this._parseRels(entry);
            break;
          case 'xl/workbook.xml':
            await this._parseWorkbook(entry);
            break;
          case 'xl/sharedStrings.xml':
            yield* this._parseSharedStrings(entry);
            break;
          case 'xl/styles.xml':
            await this._parseStyles(entry);
            break;
          default:
            if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
              match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
              sheetNo = match![1];
              if (this.sharedStrings && this.workbookRels) {
                yield* this._parseWorksheet(entry, sheetNo);
              } else {
                // Spool worksheet XML to temp until shared strings / rels are available
                if (!tempDir) {
                  tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'exceljs-'));
                }
                const filePath = path.join(tempDir, `sheet${sheetNo}.xml`);
                const tempStream = fs.createWriteStream(filePath);
                await entry.pipeTo(tempStream);
                waitingWorkSheets.push({sheetNo, path: filePath});
              }
            } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
              match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
              sheetNo = match![1];
              yield* this._parseHyperlinks(entry, sheetNo);
            }
            break;
        }
      }

      for (const {sheetNo, path: sheetPath} of waitingWorkSheets) {
        const fileStream = fs.createReadStream(sheetPath);
        yield* this._parseWorksheet(fileStream, sheetNo);
      }
    } finally {
      if (tempDir) {
        await fs.promises.rm(tempDir, {recursive: true, force: true}).catch(() => {});
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _emitEntry(payload: Record<string, any>): void {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async _parseRels(entry: any): Promise<void> {
    const xform = new RelationshipsXform();
    this.workbookRels = (await xform.parseStream(entry)) ?? undefined;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async _parseWorkbook(entry: any): Promise<void> {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(entry);

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async *_parseSharedStrings(entry: any): AsyncGenerator<any> {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text: string | null = null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let richText: any[] = [];
    let index = 0;
    let font: FontState = null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const emitted: any[] = [];
    for await (const _chunk of eachSaxChunk(entry, {
      onOpen: (name, attr) => {
        switch (name) {
          case 'b':
            font = font || {};
            font.bold = true;
            break;
          case 'charset':
            font = font || {};
            font.charset = parseInt(attr('charset') || '', 10);
            break;
          case 'color':
            font = font || {};
            font.color = {};
            if (attr('rgb')) {
              font.color.argb = attr('argb');
            }
            if (attr('val')) {
              font.color.argb = attr('val');
            }
            if (attr('theme')) {
              font.color.theme = attr('theme');
            }
            break;
          case 'family':
            font = font || {};
            font.family = parseInt(attr('val') || '', 10);
            break;
          case 'i':
            font = font || {};
            font.italic = true;
            break;
          case 'outline':
            font = font || {};
            font.outline = true;
            break;
          case 'rFont':
            font = font || {};
            font.name = undefined;
            break;
          case 'si':
            font = null;
            richText = [];
            text = null;
            break;
          case 'sz':
            font = font || {};
            font.size = parseInt(attr('val') || '', 10);
            break;
          case 'strike':
            break;
          case 't':
            text = null;
            break;
          case 'u':
            font = font || {};
            font.underline = true;
            break;
          case 'vertAlign':
            font = font || {};
            font.vertAlign = attr('val');
            break;
          default:
            break;
        }
      },
      onText: value => {
        text = text ? text + value : value;
      },
      onClose: name => {
        switch (name) {
          case 'r':
            richText.push({
              font,
              text,
            });

            font = null;
            text = null;
            break;
          case 'si':
            if (this.options.sharedStrings === 'cache') {
              this.sharedStrings!.push(richText.length ? {richText} : text);
            } else if (this.options.sharedStrings === 'emit') {
              emitted.push({index: index++, text: richText.length ? {richText} : text});
            }

            richText = [];
            font = null;
            text = null;
            break;
          default:
            break;
        }
      },
    })) {
      while (emitted.length) {
        yield emitted.shift();
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async _parseStyles(entry: any): Promise<void> {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(entry);
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  *_parseWorksheet(
    iterator: any,
    sheetNo: string | number,
  ): Generator<{eventType: string; value: any}> {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      options: this.options as any,
    });

    const matchingRel = (this.workbookRels || []).find(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (rel: any) => rel.Target === `worksheets/sheet${sheetNo}.xml`,
    );
    const matchingSheet =
      matchingRel &&
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (this.model.sheets || []).find((sheet: any) => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  *_parseHyperlinks(
    iterator: any,
    sheetNo: string | number,
  ): Generator<{eventType: string; value: any}> {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      options: this.options as any,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

export default WorkbookReader;
export {WorkbookReader};
