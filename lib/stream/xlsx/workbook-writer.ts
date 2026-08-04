import fs from 'fs';
import type {Writable} from 'node:stream';
import StreamBuf from '../../utils/stream-buf.js';
import StreamZipWriter from '../../utils/stream-zip-writer.js';
import RelType from '../../xlsx/rel-type.js';
import StylesXform from '../../xlsx/xform/style/styles-xform.js';
import SharedStrings from '../../utils/shared-strings.js';
import DefinedNames from '../../doc/defined-names.js';
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

export interface WorkbookWriterImage {
  extension: string;
  filename?: string;
  buffer?: Uint8Array | Buffer;
  base64?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

export interface WorkbookWriterMedia extends WorkbookWriterImage {
  type: string;
  name: string;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type WorksheetWriterInstance = any;

class WorkbookWriter {
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted: Date | undefined;
  useSharedStrings: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sharedStrings: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  styles: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _definedNames: any;
  _worksheets: (WorksheetWriterInstance | undefined)[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  views: any[];
  zipOptions: Partial<ZipWriterOptions> | undefined;
  media: WorkbookWriterMedia[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  commentRefs: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  zip: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  stream: any;
  promise: Promise<unknown[]>;

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

    // defined names
    this._definedNames = new DefinedNames();

    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;

    this.media = [];
    this.commentRefs = [];

    this.zip = new StreamZipWriter(this.zipOptions);
    if (options.stream) {
      this.stream = options.stream;
    } else if (options.filename) {
      this.stream = fs.createWriteStream(options.filename);
    } else {
      this.stream = new StreamBuf();
    }
    this.zip.pipe(this.stream);

    // these bits can be added right now
    this.promise = Promise.all([this.addThemes(), this.addOfficeRels()]);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  get definedNames(): any {
    return this._definedNames;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _openStream(path: string): any {
    const stream = new StreamBuf({bufSize: 65536, batch: true});
    this.zip.append(stream, {name: path});
    stream.on('finish', () => {
      stream.emit('zipped');
    });
    return stream;
  }

  _commitWorksheets(): Promise<unknown> {
    const commitWorksheet = function (worksheet: WorksheetWriterInstance) {
      if (!worksheet.committed) {
        return new Promise<void>(resolve => {
          worksheet.stream.on('zipped', () => {
            resolve();
          });
          worksheet.commit();
        });
      }
      return Promise.resolve();
    };
    // if there are any uncommitted worksheets, commit them now and wait
    const promises = this._worksheets.map(commitWorksheet);
    if (promises.length) {
      return Promise.all(promises);
    }
    return Promise.resolve();
  }

  async commit(): Promise<this> {
    // commit all worksheets, then add suplimentary files
    await this.promise;
    await this.addMedia();
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

  addImage(image: WorkbookWriterImage): number {
    const id = this.media.length;
    const medium = Object.assign({}, image, {
      type: 'image',
      name: `image${id}.${image.extension}`,
    }) as WorkbookWriterMedia;
    this.media.push(medium);
    return id;
  }

  getImage(id: number): WorkbookWriterMedia {
    return this.media[id];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  addWorksheet(name?: string, options?: any): WorksheetWriterInstance {
    // it's possible to add a worksheet with different than default
    // shared string handling
    // in fact, it's even possible to switch it mid-sheet
    options = options || {};
    const useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

    if (options.tabColor) {
      // oxlint-disable-next-line no-console
      console.trace('tabColor option has moved to { properties: tabColor: {...} }');
      options.properties = Object.assign(
        {
          tabColor: options.tabColor,
        },
        options.properties,
      );
    }

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

  getWorksheet(id?: number | string): WorksheetWriterInstance | undefined {
    if (id === undefined) {
      return this._worksheets.find(() => true);
    }
    if (typeof id === 'number') {
      return this._worksheets[id];
    }
    if (typeof id === 'string') {
      return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
    }
    return undefined;
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
        commentRefs: this.commentRefs,
        media: this.media,
      };
      const xform = new ContentTypesXform();
      const xml = xform.toXml(model);
      this.zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  }

  addMedia(): Promise<unknown[]> {
    return Promise.all(
      this.media.map(medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}`;
          if (medium.filename) {
            return this.zip.file(medium.filename, {name: filename});
          }
          if (medium.buffer) {
            return this.zip.append(medium.buffer, {name: filename});
          }
          if (medium.base64) {
            const dataimg64 = medium.base64;
            const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            return this.zip.append(content, {name: filename, base64: true});
          }
        }
        throw new Error('Unsupported media');
      }),
    );
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
      definedNames: this._definedNames.model,
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

  _finalize(): Promise<this> {
    return new Promise((resolve, reject) => {
      this.stream.on('error', reject);
      this.stream.on('finish', () => {
        resolve(this);
      });
      this.zip.on('error', reject);

      this.zip.finalize();
    });
  }
}

export default WorkbookWriter;
export {WorkbookWriter};
