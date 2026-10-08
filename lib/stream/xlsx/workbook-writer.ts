import fs from 'fs';
import type {Writable} from 'node:stream';
import {finished} from 'node:stream/promises';
import StreamZipWriter from '../../utils/stream-zip-writer.js';
import RelType from '../../xlsx/rel-type.js';
import StylesXform from '../../xlsx/xform/style/styles-xform.js';
import SharedStrings from '../../utils/shared-strings.js';
import CoreXform from '../../xlsx/xform/core/core-xform.js';
import RelationshipsXform from '../../xlsx/xform/core/relationships-xform.js';
import ContentTypesXform from '../../xlsx/xform/core/content-types-xform.js';
import AppXform from '../../xlsx/xform/core/app-xform.js';
import WorkbookXform from '../../xlsx/xform/book/workbook-xform.js';
import SharedStringsXform from '../../xlsx/xform/strings/shared-strings-xform.js';
import WorksheetWriter from './worksheet-writer.js';
import theme1Xml from '../../xlsx/xml/theme1.js';

export interface ZipWriterOptions {
  level?: number;
  store?: boolean;
  compression?: 'STORE' | 'DEFLATE' | string;
  zlib?: {level?: number};
}

export interface WorkbookWriterOptions {
  stream?: Writable;
  filename?: string;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: Partial<ZipWriterOptions>;
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type WorksheetWriterInstance = WorksheetWriter;

class WorkbookWriter {
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted: Date | undefined;
  useSharedStrings: boolean;
  sharedStrings: SharedStrings;
  styles: StylesXform;
  _worksheets: (WorksheetWriterInstance | undefined)[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  views: any[];
  zipOptions: Partial<ZipWriterOptions> | undefined;
  zip: StreamZipWriter;
  stream: Writable;
  promise: Promise<unknown[]>;
  private completion: Promise<void>;

  constructor(options?: Partial<WorkbookWriterOptions>) {
    options = options || {};

    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator || 'ExcelJS';
    this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
    this.lastPrinted = options.lastPrinted;

    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;
    this.sharedStrings = new SharedStrings();

    // style manager
    this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock();

    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;

    this.zip = new StreamZipWriter(this.zipOptions);
    if (options.stream) {
      this.stream = options.stream;
    } else if (options.filename) {
      this.stream = fs.createWriteStream(options.filename);
    } else {
      throw new Error('WorkbookWriter requires a stream or filename');
    }
    this.completion = finished(this.stream, {cleanup: true, readable: false});
    this.completion.catch(() => {});
    // Observe errors before any metadata or worksheet writes start.
    this.zip.on('error', () => {});
    this.zip.pipe(this.stream);

    // these bits can be added right now
    this.promise = Promise.all([this.addThemes(), this.addOfficeRels()]);
    this.promise.catch(() => {});
  }

  _openStream(path: string): Writable {
    return this.zip.openEntry(path);
  }

  async _commitWorksheets(): Promise<void> {
    await Promise.all(
      this._worksheets
        .filter((sheet): sheet is WorksheetWriter => !!sheet)
        .map(sheet => {
          const completion = finished(sheet.stream, {cleanup: true});
          if (!sheet.committed) sheet.commit();
          return completion;
        }),
    );
  }

  abort(error: Error): void {
    this.zip.abort(error);
  }

  async commit(): Promise<this> {
    // commit all worksheets, then add suplimentary files
    await this.promise;
    await this._commitWorksheets();
    await Promise.all([
      this.addContentTypes(),
      this.addApp(),
      this.addCore(),
      this.addSharedStrings(),
      this.addStyles(),
      this.addWorkbookRels(),
    ]);
    await this.addWorkbook();
    return this._finalize();
  }

  get nextId(): number {
    // find the next unique spot to add worksheet
    let i: number;
    for (i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  addWorksheet(name?: string, options?: any): WorksheetWriterInstance {
    // it's possible to add a worksheet with different than default
    // shared string handling
    // in fact, it's even possible to switch it mid-sheet
    options = options || {};
    const useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

    const id = this.nextId;
    name = name || `sheet${id}`;

    const worksheet = new WorksheetWriter({
      id,
      name,
      workbook: this,
      useSharedStrings,
      properties: options.properties,
      state: options.state,
      pageSetup: options.pageSetup,
      views: options.views,
      autoFilter: options.autoFilter,
      headerFooter: options.headerFooter,
    });

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  addStyles(): Promise<void> {
    return new Promise(resolve => {
      this.zip.append(this.styles.xml, {name: 'xl/styles.xml'});
      resolve();
    });
  }

  addThemes(): Promise<void> {
    return new Promise(resolve => {
      this.zip.append(theme1Xml, {name: 'xl/theme/theme1.xml'});
      resolve();
    });
  }

  addOfficeRels(): Promise<void> {
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        {Id: 'rId1', Type: RelType.OfficeDocument, Target: 'xl/workbook.xml'},
        {Id: 'rId2', Type: RelType.CoreProperties, Target: 'docProps/core.xml'},
        {Id: 'rId3', Type: RelType.ExtenderProperties, Target: 'docProps/app.xml'},
      ]);
      this.zip.append(xml, {name: '/_rels/.rels'});
      resolve();
    });
  }

  addContentTypes(): Promise<void> {
    return new Promise(resolve => {
      const model = {
        worksheets: this._worksheets.filter(Boolean),
        sharedStrings: this.sharedStrings,
        commentRefs: [],
        media: [],
      };
      const xform = new ContentTypesXform();
      const xml = xform.toXml(model);
      this.zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  }

  addApp(): Promise<void> {
    return new Promise(resolve => {
      const model = {
        worksheets: this._worksheets.filter(Boolean),
      };
      const xform = new AppXform();
      const xml = xform.toXml(model);
      this.zip.append(xml, {name: 'docProps/app.xml'});
      resolve();
    });
  }

  addCore(): Promise<void> {
    return new Promise(resolve => {
      const coreXform = new CoreXform();
      const xml = coreXform.toXml(this);
      this.zip.append(xml, {name: 'docProps/core.xml'});
      resolve();
    });
  }

  addSharedStrings(): Promise<void> {
    if (this.sharedStrings.count) {
      return new Promise(resolve => {
        const sharedStringsXform = new SharedStringsXform();
        const xml = sharedStringsXform.toXml(this.sharedStrings);
        this.zip.append(xml, {name: '/xl/sharedStrings.xml'});
        resolve();
      });
    }
    return Promise.resolve();
  }

  addWorkbookRels(): Promise<void> {
    let count = 1;
    const relationships: Array<{Id: string; Type: string; Target: string}> = [
      {Id: `rId${count++}`, Type: RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (this.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    this._worksheets.forEach(worksheet => {
      if (worksheet) {
        worksheet.rId = `rId${count++}`;
        relationships.push({
          Id: worksheet.rId,
          Type: RelType.Worksheet,
          Target: `worksheets/sheet${worksheet.id}.xml`,
        });
      }
    });
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml(relationships);
      this.zip.append(xml, {name: '/xl/_rels/workbook.xml.rels'});
      resolve();
    });
  }

  addWorkbook(): Promise<void> {
    const {zip} = this;
    const model = {
      worksheets: this._worksheets.filter(Boolean),
      definedNames: [],
      views: this.views,
      properties: {},
      calcProperties: {},
    };

    return new Promise(resolve => {
      const xform = new WorkbookXform();
      xform.prepare(model);
      zip.append(xform.toXml(model), {name: '/xl/workbook.xml'});
      resolve();
    });
  }

  async _finalize(): Promise<this> {
    this.zip.finalize();
    await this.completion;
    return this;
  }
}

export default WorkbookWriter;
export {WorkbookWriter};
