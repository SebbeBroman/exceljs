import fs from 'fs';
import BufferZipWriter, {resolveZipLevel} from '../utils/buffer-zip.js';
import XmlStream from '../utils/xml-stream.js';
import MinimalStyles from './xform/style/minimal-styles.js';
import CoreXform from './xform/core/core-xform.js';
import SharedStringsXform from './xform/strings/shared-strings-xform.js';
import RelationshipsXform from './xform/core/relationships-xform.js';
import ContentTypesXform from './xform/core/content-types-xform.js';
import AppXform from './xform/core/app-xform.js';
import WorkbookXform from './xform/book/workbook-xform.js';
import WorksheetXform from './xform/sheet/worksheet-xform.js';
import {
  loadDrawingXform,
  loadTableXform,
  loadCommentsXform,
  loadVmlNotesXform,
  loadPivotXforms,
} from './lazy-xforms.js';
import RelType from './rel-type.js';

/** Workbook host that owns the model XLSX reads into / writes from */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type XlsxWorkbookHost = {model: any};

export interface XlsxWriteOptions {
  useSharedStrings?: boolean;
  useStyles?: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  zip?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type XlsxModel = any;

function sheetNeedsCf(worksheet: XlsxModel): boolean {
  return Boolean(
    worksheet && worksheet.conditionalFormattings && worksheet.conditionalFormattings.length,
  );
}

// theme1 XML is large (~8KB min); load only when writing default theme
let theme1XmlPromise: Promise<string> | undefined;
function loadTheme1Xml(): Promise<string> {
  if (!theme1XmlPromise) {
    theme1XmlPromise = import('./xml/theme1.js').then(m => m.default);
  }
  return theme1XmlPromise;
}

function fsReadFileAsync(
  filename: string,
  options?: Parameters<typeof fs.readFile>[1],
): Promise<Buffer | string> {
  return new Promise((resolve, reject) => {
    fs.readFile(
      filename,
      options as any,
      (error: NodeJS.ErrnoException | null, data: Buffer | string) => {
        if (error) {
          reject(error);
        } else {
          resolve(data);
        }
      },
    );
  });
}

class XlsxWriter {
  workbook: XlsxWorkbookHost;

  constructor(workbook: XlsxWorkbookHost) {
    this.workbook = workbook;
  }

  async addMedia(zip: any, model: XlsxModel): Promise<void> {
    await Promise.all(
      (model.media as XlsxModel[]).map(async (medium: XlsxModel) => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}.${medium.extension}`;
          if (medium.filename) {
            const data = await fsReadFileAsync(medium.filename);
            return zip.append(data, {name: filename});
          }
          if (medium.buffer) {
            return zip.append(medium.buffer, {name: filename});
          }
          if (medium.base64) {
            const dataimg64 = medium.base64;
            const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            return zip.append(content, {name: filename, base64: true});
          }
        }
        throw new Error('Unsupported media');
      }),
    );
  }

  async addDrawings(zip: any, model: XlsxModel): Promise<void> {
    const hasDrawing = (model.worksheets as XlsxModel[]).some((ws: XlsxModel) => ws.drawing);
    if (!hasDrawing) return;

    const DrawingXform = await loadDrawingXform();
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
      const {drawing} = worksheet;
      if (drawing) {
        drawingXform.prepare(drawing, {});
        let xml = drawingXform.toXml(drawing);
        zip.append(xml, {name: `xl/drawings/${drawing.name}.xml`});

        xml = relsXform.toXml(drawing.rels);
        zip.append(xml, {name: `xl/drawings/_rels/${drawing.name}.xml.rels`});
      }
    });
  }

  async addTables(zip: any, model: XlsxModel): Promise<void> {
    const hasTables = (model.worksheets as XlsxModel[]).some(
      (ws: XlsxModel) => ws.tables && ws.tables.length,
    );
    if (!hasTables) return;

    const TableXform = await loadTableXform();
    const tableXform = new TableXform();

    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
      const {tables} = worksheet;
      (tables as XlsxModel[]).forEach((table: XlsxModel) => {
        tableXform.prepare(table, {});
        const tableXml = tableXform.toXml(table);
        zip.append(tableXml, {name: `xl/tables/${table.target}`});
      });
    });
  }

  async addPivotTables(zip: any, model: XlsxModel): Promise<void> {
    if (!model.pivotTables.length) return;

    const pivotTable = model.pivotTables[0];
    const {PivotCacheRecordsXform, PivotCacheDefinitionXform, PivotTableXform} =
      await loadPivotXforms();

    const pivotCacheRecordsXform = new PivotCacheRecordsXform();
    const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
    const pivotTableXform = new PivotTableXform();
    const relsXform = new RelationshipsXform();

    // pivot cache records
    // --------------------------------------------------
    // copy of the source data.
    //
    // Note: cells in the columns of the source data which are part
    // of the "rows" or "columns" of the pivot table configuration are
    // replaced by references to their __cache field__ identifiers.
    // See "pivot cache definition" below.

    let xml = pivotCacheRecordsXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotCache/pivotCacheRecords1.xml'});

    // pivot cache definition
    // --------------------------------------------------
    // cache source (source data):
    //    ref="A1:E7" on sheet="Sheet1"
    // cache fields:
    //    - 0: "A" (a1, a2, a3)
    //    - 1: "B" (b1, b2)
    //    - ...

    xml = pivotCacheDefinitionXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotCache/pivotCacheDefinition1.xml'});

    xml = relsXform.toXml([
      {
        Id: 'rId1',
        Type: RelType.PivotCacheRecords,
        Target: 'pivotCacheRecords1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'});

    // pivot tables (on destination worksheet)
    // --------------------------------------------------
    // location: ref="A3:E15"
    // pivotFields
    // rowFields and rowItems
    // colFields and colItems
    // dataFields
    // pivotTableStyleInfo

    xml = pivotTableXform.toXml(pivotTable);
    zip.append(xml, {name: 'xl/pivotTables/pivotTable1.xml'});

    xml = relsXform.toXml([
      {
        Id: 'rId1',
        Type: RelType.PivotCacheDefinition,
        Target: '../pivotCache/pivotCacheDefinition1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotTables/_rels/pivotTable1.xml.rels'});
  }

  async addContentTypes(zip: any, model: XlsxModel): Promise<void> {
    const xform = new ContentTypesXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: '[Content_Types].xml'});
  }

  async addApp(zip: any, model: XlsxModel): Promise<void> {
    const xform = new AppXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: 'docProps/app.xml'});
  }

  async addCore(zip: any, model: XlsxModel): Promise<void> {
    const coreXform = new CoreXform();
    zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
  }

  async addThemes(zip: any, model: XlsxModel): Promise<void> {
    let themes = model.themes;
    if (!themes) {
      const theme1Xml = await loadTheme1Xml();
      themes = {theme1: theme1Xml};
    }
    Object.keys(themes).forEach(name => {
      const xml = themes[name];
      const path = `xl/theme/${name}.xml`;
      zip.append(xml, {name: path});
    });
  }

  async addOfficeRels(zip: any, _model?: XlsxModel): Promise<void> {
    const xform = new RelationshipsXform();
    const xml = xform.toXml([
      {Id: 'rId1', Type: RelType.OfficeDocument, Target: 'xl/workbook.xml'},
      {Id: 'rId2', Type: RelType.CoreProperties, Target: 'docProps/core.xml'},
      {Id: 'rId3', Type: RelType.ExtenderProperties, Target: 'docProps/app.xml'},
    ]);
    zip.append(xml, {name: '_rels/.rels'});
  }

  async addWorkbookRels(zip: any, model: XlsxModel): Promise<void> {
    let count = 1;
    const relationships: Array<{Id: string; Type: string; Target: string}> = [
      {Id: `rId${count++}`, Type: RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (model.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    if ((model.pivotTables || []).length) {
      const pivotTable = model.pivotTables[0];
      pivotTable.rId = `rId${count++}`;
      relationships.push({
        Id: pivotTable.rId,
        Type: RelType.PivotCacheDefinition,
        Target: 'pivotCache/pivotCacheDefinition1.xml',
      });
    }
    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
      worksheet.rId = `rId${count++}`;
      relationships.push({
        Id: worksheet.rId,
        Type: RelType.Worksheet,
        Target: `worksheets/sheet${worksheet.id}.xml`,
      });
    });
    const xform = new RelationshipsXform();
    const xml = xform.toXml(relationships);
    zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
  }

  async addSharedStrings(zip: any, model: XlsxModel): Promise<void> {
    if (model.sharedStrings && model.sharedStrings.count) {
      zip.append(model.sharedStrings.xml, {name: 'xl/sharedStrings.xml'});
    }
  }

  async addStyles(zip: any, model: XlsxModel): Promise<void> {
    const result = model.styles instanceof MinimalStyles ? model.styles.toXml() : model.styles.xml;
    const xml = typeof result === 'string' ? result : await result;
    if (xml) {
      zip.append(xml, {name: 'xl/styles.xml'});
    }
  }

  async addWorkbook(zip: any, model: XlsxModel): Promise<void> {
    const xform = new WorkbookXform();
    zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
  }

  async addWorksheets(zip: any, model: XlsxModel): Promise<void> {
    // preparation phase
    const worksheetXform = new WorksheetXform();
    const relationshipsXform = new RelationshipsXform();

    if ((model.worksheets || []).some(sheetNeedsCf)) {
      await worksheetXform.installCfXforms();
    }

    const hasComments = model.worksheets.some(
      (ws: XlsxModel) => ws.comments && ws.comments.length > 0,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let commentsXform: any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let vmlNotesXform: any;
    if (hasComments) {
      const CommentsXform = await loadCommentsXform();
      const VmlNotesXform = await loadVmlNotesXform();
      commentsXform = new CommentsXform();
      vmlNotesXform = new VmlNotesXform();
    }

    // write sheets
    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
      let xmlStream = new XmlStream();
      worksheetXform.render(xmlStream, worksheet);
      zip.append(xmlStream.xml, {name: `xl/worksheets/sheet${worksheet.id}.xml`});

      if (worksheet.rels && worksheet.rels.length) {
        xmlStream = new XmlStream();
        relationshipsXform.render(xmlStream, worksheet.rels);
        zip.append(xmlStream.xml, {name: `xl/worksheets/_rels/sheet${worksheet.id}.xml.rels`});
      }

      if (worksheet.comments.length > 0) {
        xmlStream = new XmlStream();
        commentsXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: `xl/comments${worksheet.id}.xml`});

        xmlStream = new XmlStream();
        vmlNotesXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: `xl/drawings/vmlDrawing${worksheet.id}.vml`});
      }
    });
  }

  async prepareModel(model: XlsxModel, options: XlsxWriteOptions = {}): Promise<void> {
    // ensure following properties have sane values
    model.creator = model.creator || 'ExcelJS';
    model.lastModifiedBy = model.lastModifiedBy || 'ExcelJS';
    model.created = model.created || new Date();
    model.modified = model.modified || new Date();

    model.useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : true;
    model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

    // Manage the shared strings
    model.sharedStrings = new SharedStringsXform();

    // add a style manager to handle cell formats, fonts, etc.
    model.styles = model.useStyles
      ? new (await import('./xform/style/styles-xform.js')).default(true)
      : new MinimalStyles();

    // prepare all of the things before the render
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform();

    if ((model.worksheets || []).some(sheetNeedsCf)) {
      await worksheetXform.installCfXforms();
    }

    workbookXform.prepare(model);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const worksheetOptions: any = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: model.properties.date1904,
      drawingsCount: 0,
      media: model.media,
    };
    worksheetOptions.drawings = model.drawings = [];
    worksheetOptions.commentRefs = model.commentRefs = [];
    let tableCount = 0;
    model.tables = [];
    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
      // assign unique filenames to tables
      const tables = worksheet.tables
        ? Array.isArray(worksheet.tables)
          ? worksheet.tables
          : Object.values(worksheet.tables)
        : [];
      (tables as XlsxModel[]).forEach((table: XlsxModel) => {
        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
        model.tables.push(table);
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  }

  /**
   * Render every package part into `zip.append`. The buffered `writeBuffer()` path collects these parts and compresses once.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async renderParts(
    zip: {append(data: unknown, options: {name: string; [key: string]: any}): unknown},
    model: XlsxModel,
  ): Promise<void> {
    // render
    await this.addContentTypes(zip, model);
    await this.addOfficeRels(zip, model);
    await this.addWorkbookRels(zip, model);
    await this.addWorksheets(zip, model);
    await this.addSharedStrings(zip, model); // always after worksheets
    await this.addDrawings(zip, model);
    await this.addTables(zip, model);
    await this.addPivotTables(zip, model);
    await Promise.all([this.addThemes(zip, model), this.addStyles(zip, model)]);
    await this.addMedia(zip, model);
    await Promise.all([this.addApp(zip, model), this.addCore(zip, model)]);
    await this.addWorkbook(zip, model);
  }

  async writeBuffer(options?: XlsxWriteOptions): Promise<unknown> {
    options = options || {};
    const model = this.workbook.model;
    await this.prepareModel(model, options);
    // Buffered fast path: every part is already materialized (XML strings,
    // media buffers), so collect + single synchronous zipSync instead of the
    // streaming ZipWriter's event-loop hops. Same parts, same level.
    const zip = new BufferZipWriter();
    await this.renderParts(zip, model);
    return zip.toBytes(resolveZipLevel(options.zip));
  }
}

export default XlsxWriter;
