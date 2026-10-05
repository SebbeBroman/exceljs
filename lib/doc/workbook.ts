import Worksheet from './worksheet.js';
import DefinedNames from './defined-names.js';
import XLSX from '../xlsx/xlsx.js';
import type {
  AddWorksheetOptions,
  CalculationProperties,
  Color,
  Image as PublicImage,
  Media,
  WorkbookModel,
  WorkbookProperties,
  WorkbookView,
} from '../../index.js';
import type {PivotTable} from './pivot-table.js';
import type {WorksheetModelData, WorksheetOptions} from './worksheet.js';

// INTERNAL Doc Workbook (mutable class model). Used by the materialize/load
// bridge and historical tests — not part of the public package API.
// Public CSV: named `csv` from `@sebbebroman/exceljs`. Tests enable class CSV via lib/csv-entry.ts.

/**
 * Constructor type for optional CSV module.
 * Parameter/return typed loosely so the csv entry can assign without circular
 * CsvWorkbook ↔ Workbook constraints.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type CsvConstructor = new (workbook: any) => any;

class Workbook {
  /** Optional CSV class; set via enableCsv from lib/csv-entry.ts (tests/bridge only). */
  static CSV: CsvConstructor | null = null;

  category: string;
  company: string;
  created: Date;
  description: string;
  keywords: string;
  manager: string;
  modified: Date;
  properties: WorkbookProperties | Record<string, unknown>;
  calcProperties: CalculationProperties | Record<string, unknown>;
  _worksheets: Array<Worksheet | undefined>;
  subject: string;
  title: string;
  views: WorkbookView[];
  media: Array<Media | (PublicImage & {type?: string})>;
  pivotTables: PivotTable[];
  _definedNames: DefinedNames;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  language?: string;
  revision?: Date | string | number;
  contentStatus?: string;
  _themes?: string[] | unknown;
  private _xlsx?: InstanceType<typeof XLSX>;
  private _csv?: InstanceType<CsvConstructor>;

  constructor() {
    this.category = '';
    this.company = '';
    this.created = new Date();
    this.description = '';
    this.keywords = '';
    this.manager = '';
    this.modified = this.created;
    this.properties = {};
    this.calcProperties = {};
    this._worksheets = [];
    this.subject = '';
    this.title = '';
    this.views = [];
    this.media = [];
    this.pivotTables = [];
    this._definedNames = new DefinedNames();
  }

  get xlsx(): InstanceType<typeof XLSX> {
    if (!this._xlsx) this._xlsx = new XLSX(this);
    return this._xlsx;
  }

  get csv(): InstanceType<CsvConstructor> {
    if (!Workbook.CSV) {
      throw new Error(
        'CSV support is not loaded on Doc Workbook. For the public API use `csv` from @sebbebroman/exceljs. For tests, import lib/csv-entry (enableCsv).',
      );
    }
    if (!this._csv) this._csv = new Workbook.CSV(this);
    return this._csv;
  }

  get nextId(): number {
    // find the next unique spot to add worksheet
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  addWorksheet(
    name?: string,
    options?:
      | Partial<AddWorksheetOptions>
      | string
      | Partial<Color>
      | (Partial<Color> & Record<string, unknown>),
  ): Worksheet {
    const id = this.nextId;

    // if options is a color, call it tabColor (and signal deprecated message)
    let opts: Partial<AddWorksheetOptions> | undefined | null = options as
      | Partial<AddWorksheetOptions>
      | undefined
      | null;
    if (options) {
      if (typeof options === 'string') {
        // oxlint-disable-next-line no-console
        console.trace(
          'tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { argb: "rbg value" } }',
        );
        opts = {
          properties: {
            tabColor: {argb: options},
          },
        };
      } else if (
        (options as Partial<Color>).argb ||
        (options as Partial<Color>).theme !== undefined ||
        (options as Partial<Color> & {indexed?: number}).indexed !== undefined
      ) {
        // oxlint-disable-next-line no-console
        console.trace(
          'tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }',
        );
        opts = {
          properties: {
            tabColor: options as Partial<Color>,
          },
        };
      }
    }

    const lastOrderNo = this._worksheets.reduce(
      (acc, ws) => (ws && ws.orderNo > acc ? ws.orderNo : acc),
      0,
    );
    const worksheetOptions = Object.assign({}, opts, {
      id,
      name,
      orderNo: lastOrderNo + 1,
      workbook: this,
    }) as WorksheetOptions;

    const worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  removeWorksheetEx(worksheet: Worksheet): void {
    delete this._worksheets[worksheet.id];
  }

  removeWorksheet(id?: number | string): void {
    const worksheet = this.getWorksheet(id);
    if (worksheet) {
      worksheet.destroy();
    }
  }

  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) {
      return this._worksheets.find(Boolean);
    }
    if (typeof id === 'number') {
      return this._worksheets[id];
    }
    if (typeof id === 'string') {
      return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
    }
    return undefined;
  }

  get worksheets(): Worksheet[] {
    // return a clone of _worksheets
    return this._worksheets
      .slice(1)
      .sort((a, b) => a!.orderNo - b!.orderNo)
      .filter(Boolean) as Worksheet[];
  }

  eachSheet(iteratee: (sheet: Worksheet, id: number) => void): void {
    this.worksheets.forEach(sheet => {
      iteratee(sheet, sheet.id);
    });
  }

  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  clearThemes(): void {
    // Note: themes are not an exposed feature, meddle at your peril!
    this._themes = undefined;
  }

  addImage(image: PublicImage): number {
    // TODO:  validation?
    const id = this.media.length;
    this.media.push(Object.assign({}, image, {type: 'image'}));
    return id;
  }

  getImage(id: number): Media | (PublicImage & {type?: string}) {
    return this.media[id];
  }

  get model(): ReturnType<Workbook['getXlsxModel']> {
    const model = this.getXlsxModel();
    // Keep the legacy snapshot's two independent worksheet model arrays.
    model.sheets = this.worksheets.map(ws => ws.model).filter(Boolean) as never;
    return model;
  }

  /** @internal Write once per sheet; XLSX consumes worksheets, not a second cell graph. */
  getXlsxModel(): WorkbookModel & {
    worksheets: WorksheetModelData[];
    sheets: WorksheetModelData[];
    pivotTables: PivotTable[];
    calcProperties: CalculationProperties | Record<string, unknown>;
    themes?: unknown;
  } {
    const worksheets = this.worksheets.map(worksheet => worksheet.model);
    return {
      creator: this.creator || 'Unknown',
      lastModifiedBy: this.lastModifiedBy || 'Unknown',
      lastPrinted: this.lastPrinted as Date,
      created: this.created,
      modified: this.modified,
      properties: this.properties as WorkbookProperties,
      worksheets: worksheets as never,
      sheets: worksheets.filter(Boolean) as never,
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category: this.category,
      description: this.description,
      language: this.language as string,
      revision: this.revision as Date,
      contentStatus: this.contentStatus as string,
      themes: this._themes as string[],
      media: this.media as Media[],
      pivotTables: this.pivotTables,
      calcProperties: this.calcProperties,
    };
  }

  set model(
    value: Omit<WorkbookModel, 'worksheets' | 'sheets'> & {
      worksheets: WorksheetModelData[];
      sheets?: Array<{id?: number}>;
      pivotTables?: PivotTable[];
      themes?: unknown;
      calcProperties?: CalculationProperties | Record<string, unknown>;
    },
  ) {
    this.creator = value.creator;
    this.lastModifiedBy = value.lastModifiedBy;
    this.lastPrinted = value.lastPrinted;
    this.created = value.created;
    this.modified = value.modified;
    this.company = value.company;
    this.manager = value.manager;
    this.title = value.title;
    this.subject = value.subject;
    this.keywords = value.keywords;
    this.category = value.category;
    this.description = value.description;
    this.language = value.language;
    this.revision = value.revision;
    this.contentStatus = value.contentStatus;

    this.properties = value.properties;
    this.calcProperties = value.calcProperties || {};
    this._worksheets = [];
    value.worksheets.forEach(worksheetModel => {
      const {id, name, state} = worksheetModel;
      const orderNo = value.sheets && value.sheets.findIndex(ws => ws.id === id);
      const worksheet = new Worksheet({
        id,
        name,
        orderNo: orderNo as number,
        state,
        workbook: this,
      });
      this._worksheets[id as number] = worksheet;
      worksheet.model = worksheetModel;
    });

    this._definedNames.model = value.definedNames;
    this.views = value.views;
    this._themes = value.themes;
    this.media = (value.media || []) as Array<Media | (PublicImage & {type?: string})>;
    this.pivotTables = value.pivotTables || [];
  }
}

export default Workbook;
export {Workbook};
