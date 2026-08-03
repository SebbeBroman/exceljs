import fs from 'fs';
import {fromReadable, once, stringChunks} from '../utils/async-iterator.js';
import {entryToBuffer, entryToString, unzipToFiles} from '../utils/zip-reader.js';
import ZipStream from '../utils/zip-stream.js';
import StreamBuf from '../utils/stream-buf.js';
import utils from '../utils/utils.js';
import XmlStream from '../utils/xml-stream.js';
import {
  asUint8Array,
  concat,
  from as bytesFrom,
  isBytes,
  toPublic,
  toString as bytesToString,
} from '../utils/bytes.js';
import StylesXform from './xform/style/styles-xform.js';
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
import {ensureDocFeatures} from '../doc/doc-features.js';
import __esm_0 from './rel-type.js';

function sheetNeedsCf(worksheet) {
  return Boolean(
    worksheet && worksheet.conditionalFormattings && worksheet.conditionalFormattings.length,
  );
}

function modelNeedsDocFeatures(model) {
  if (!model) return false;
  if (model.media && model.media.length) return true;
  if (model.pivotTables && model.pivotTables.length) return true;
  if (model.worksheets) {
    for (const ws of model.worksheets) {
      if (ws.media && ws.media.length) return true;
      if (ws.tables && ws.tables.length) return true;
      if (ws.pivotTables && ws.pivotTables.length) return true;
      // media on sheet model after reconcile is ws.media; before write it's from doc model
      if (ws._media && ws._media.length) return true;
    }
  }
  // tables collated onto model during prepare
  if (model.tables && model.tables.length) return true;
  return false;
}

// theme1 XML is large (~8KB min); load only when writing default theme
let theme1XmlPromise;
function loadTheme1Xml() {
  if (!theme1XmlPromise) {
    theme1XmlPromise = import('./xml/theme1.js').then(m => m.default);
  }
  return theme1XmlPromise;
}

function fsReadFileAsync(filename, options) {
  return new Promise((resolve, reject) => {
    fs.readFile(filename, options, (error, data) => {
      if (error) {
        reject(error);
      } else {
        resolve(data);
      }
    });
  });
}

class XLSX {
  constructor(workbook) {
    this.workbook = workbook;
  }

  // ===============================================================================
  // Workbook
  // =========================================================================
  // Read

  async readFile(filename, options) {
    if (!(await utils.fs.exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    try {
      const workbook = await this.read(stream, options);
      stream.close();
      return workbook;
    } catch (error) {
      stream.close();
      throw error;
    }
  }

  parseRels(stream) {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  }

  parseWorkbook(stream) {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  }

  parseSharedStrings(stream) {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  }

  async reconcile(model, options) {
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform(options);

    // Load CF xforms if any sheet has conditional formatting (after parse they may)
    const needsCf = (model.worksheets || []).some(sheetNeedsCf);
    if (needsCf) {
      await worksheetXform.installCfXforms();
    }

    // Image/Table classes needed when hydrating worksheet doc models (workbook.model setter)
    if (
      modelNeedsDocFeatures(model) ||
      (model.worksheets || []).some(
        ws => (ws.tables && ws.tables.length) || (ws.media && ws.media.length),
      )
    ) {
      await ensureDocFeatures();
    }

    workbookXform.reconcile(model);

    // reconcile drawings with their rels (lazy-load drawing xform only if present)
    const drawingNames = Object.keys(model.drawings);
    if (drawingNames.length) {
      const DrawingXform = await loadDrawingXform();
      const drawingXform = new DrawingXform();
      const drawingOptions = {
        media: model.media,
        mediaIndex: model.mediaIndex,
      };
      drawingNames.forEach(name => {
        const drawing = model.drawings[name];
        const drawingRel = model.drawingRels[name];
        if (drawingRel) {
          drawingOptions.rels = drawingRel.reduce((o, rel) => {
            o[rel.Id] = rel;
            return o;
          }, {});
          (drawing.anchors || []).forEach(anchor => {
            const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
            if (hyperlinks && drawingOptions.rels[hyperlinks.rId]) {
              hyperlinks.hyperlink = drawingOptions.rels[hyperlinks.rId].Target;
              delete hyperlinks.rId;
            }
          });
          drawingXform.reconcile(drawing, drawingOptions);
        }
      });
    }

    // reconcile tables with the default styles
    const tables = Object.values(model.tables);
    if (tables.length) {
      const TableXform = await loadTableXform();
      const tableXform = new TableXform();
      const tableOptions = {
        styles: model.styles,
      };
      tables.forEach(table => {
        tableXform.reconcile(table, tableOptions);
      });
    }

    const sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties && model.properties.date1904,
      drawings: model.drawings,
      comments: model.comments,
      tables: model.tables,
      vmlDrawings: model.vmlDrawings,
    };
    model.worksheets.forEach(worksheet => {
      worksheet.relationships = model.worksheetRels[worksheet.sheetNo];
      worksheetXform.reconcile(worksheet, sheetOptions);
    });

    // delete unnecessary parts
    delete model.worksheetHash;
    delete model.worksheetRels;
    delete model.globalRels;
    delete model.sharedStrings;
    delete model.workbookRels;
    delete model.sheetDefs;
    delete model.styles;
    delete model.mediaIndex;
    delete model.drawings;
    delete model.drawingRels;
    delete model.vmlDrawings;
  }

  async _processWorksheetEntry(stream, model, sheetNo, options, path) {
    const xform = new WorksheetXform(options);
    // Always install CF parsers for read — sheets may contain conditionalFormatting/extLst
    await xform.installCfXforms();
    const worksheet = await xform.parseStream(stream);
    worksheet.sheetNo = sheetNo;
    model.worksheetHash[path] = worksheet;
    model.worksheets.push(worksheet);
  }

  async _processCommentEntry(stream, model, name) {
    const CommentsXform = await loadCommentsXform();
    const xform = new CommentsXform();
    const comments = await xform.parseStream(stream);
    model.comments[`../${name}.xml`] = comments;
  }

  async _processTableEntry(stream, model, name) {
    const TableXform = await loadTableXform();
    const xform = new TableXform();
    const table = await xform.parseStream(stream);
    model.tables[`../tables/${name}.xml`] = table;
  }

  async _processWorksheetRelsEntry(stream, model, sheetNo) {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.worksheetRels[sheetNo] = relationships;
  }

  async _processMediaEntry(entry, model, filename) {
    const lastDot = filename.lastIndexOf('.');
    // if we can't determine extension, ignore it
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      const chunks = [];
      for await (const chunk of fromReadable(entry)) {
        chunks.push(isBytes(chunk) ? chunk : bytesFrom(chunk));
      }
      model.mediaIndex[filename] = model.media.length;
      model.mediaIndex[name] = model.media.length;
      model.media.push({
        type: 'image',
        name,
        extension,
        buffer: toPublic(concat(chunks)),
      });
    }
  }

  async _processDrawingEntry(entry, model, name) {
    const DrawingXform = await loadDrawingXform();
    const xform = new DrawingXform();
    const drawing = await xform.parseStream(entry);
    model.drawings[name] = drawing;
  }

  async _processDrawingRelsEntry(entry, model, name) {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.drawingRels[name] = relationships;
  }

  async _processVmlDrawingEntry(entry, model, name) {
    const VmlNotesXform = await loadVmlNotesXform();
    const xform = new VmlNotesXform();
    const vmlDrawing = await xform.parseStream(entry);
    model.vmlDrawings[`../drawings/${name}.vml`] = vmlDrawing;
  }

  async _processThemeEntry(entry, model, name) {
    const chunks = [];
    for await (const chunk of fromReadable(entry)) {
      chunks.push(isBytes(chunk) ? chunk : bytesFrom(chunk));
    }
    model.themes[name] = bytesToString(concat(chunks));
  }

  /**
   * @deprecated since version 4.0. You should use `#read` instead. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md
   */
  createInputStream() {
    throw new Error(
      '`XLSX#createInputStream` is deprecated. You should use `XLSX#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md',
    );
  }

  async read(stream, options) {
    const chunks = [];
    for await (const chunk of fromReadable(stream)) {
      chunks.push(chunk);
    }
    // Normalize to Uint8Array (accepts Buffer / Uint8Array from streams)
    const parts = chunks.map(c => (isBytes(c) ? c : bytesFrom(c)));
    return this.load(concat(parts), options);
  }

  async load(data, options) {
    let buffer;
    if (options && options.base64) {
      buffer = bytesFrom(data.toString(), 'base64');
    } else if (isBytes(data)) {
      buffer = data;
    } else if (data instanceof ArrayBuffer) {
      buffer = asUint8Array(data);
    } else if (ArrayBuffer.isView(data)) {
      buffer = asUint8Array(data);
    } else {
      buffer = data;
    }

    const model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {},
      comments: {},
      tables: {},
      vmlDrawings: {},
    };

    const zipFiles = await unzipToFiles(buffer);
    for (const [rawName, entryBytes] of Object.entries(zipFiles)) {
      // fflate omits pure directory entries; still skip trailing-slash keys if present
      if (!rawName || rawName.endsWith('/')) {
        continue;
      }
      const entryName = rawName[0] === '/' ? rawName.slice(1) : rawName;
      // Async iterables of chunks for SAX parsers
      let xmlSource;
      if (
        entryName.match(/xl\/media\//) ||
        // themes are stored as raw bytes then stringified later
        entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/)
      ) {
        xmlSource = once(entryToBuffer(entryBytes));
      } else {
        xmlSource = stringChunks(entryToString(entryBytes));
      }
      switch (entryName) {
        case '_rels/.rels':
          model.globalRels = await this.parseRels(xmlSource);
          break;

        case 'xl/workbook.xml': {
          const workbook = await this.parseWorkbook(xmlSource);
          model.sheets = workbook.sheets;
          model.definedNames = workbook.definedNames;
          model.views = workbook.views;
          model.properties = workbook.properties;
          model.calcProperties = workbook.calcProperties;
          break;
        }

        case 'xl/_rels/workbook.xml.rels':
          model.workbookRels = await this.parseRels(xmlSource);
          break;

        case 'xl/sharedStrings.xml':
          model.sharedStrings = new SharedStringsXform();
          await model.sharedStrings.parseStream(xmlSource);
          break;

        case 'xl/styles.xml':
          model.styles = new StylesXform();
          await model.styles.parseStream(xmlSource);
          break;

        case 'docProps/app.xml': {
          const appXform = new AppXform();
          const appProperties = await appXform.parseStream(xmlSource);
          model.company = appProperties.company;
          model.manager = appProperties.manager;
          break;
        }

        case 'docProps/core.xml': {
          const coreXform = new CoreXform();
          const coreProperties = await coreXform.parseStream(xmlSource);
          Object.assign(model, coreProperties);
          break;
        }

        default: {
          let match = entryName.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
          if (match) {
            await this._processWorksheetEntry(xmlSource, model, match[1], options, entryName);
            break;
          }
          match = entryName.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
          if (match) {
            await this._processWorksheetRelsEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
          if (match) {
            await this._processThemeEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
          if (match) {
            await this._processMediaEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
          if (match) {
            await this._processDrawingEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/(comments\d+)[.]xml/);
          if (match) {
            await this._processCommentEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/tables\/(table\d+)[.]xml/);
          if (match) {
            await this._processTableEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
          if (match) {
            await this._processDrawingRelsEntry(xmlSource, model, match[1]);
            break;
          }
          match = entryName.match(/xl\/drawings\/(vmlDrawing\d+)[.]vml/);
          if (match) {
            await this._processVmlDrawingEntry(xmlSource, model, match[1]);
            break;
          }
        }
      }
    }

    await this.reconcile(model, options);

    // apply model
    this.workbook.model = model;
    return this.workbook;
  }

  // =========================================================================
  // Write

  async addMedia(zip, model) {
    await Promise.all(
      model.media.map(async medium => {
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

  async addDrawings(zip, model) {
    const hasDrawing = model.worksheets.some(ws => ws.drawing);
    if (!hasDrawing) return;

    const DrawingXform = await loadDrawingXform();
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    model.worksheets.forEach(worksheet => {
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

  async addTables(zip, model) {
    const hasTables = model.worksheets.some(ws => ws.tables && ws.tables.length);
    if (!hasTables) return;

    const TableXform = await loadTableXform();
    const tableXform = new TableXform();

    model.worksheets.forEach(worksheet => {
      const {tables} = worksheet;
      tables.forEach(table => {
        tableXform.prepare(table, {});
        const tableXml = tableXform.toXml(table);
        zip.append(tableXml, {name: `xl/tables/${table.target}`});
      });
    });
  }

  async addPivotTables(zip, model) {
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
        Type: XLSX.RelType.PivotCacheRecords,
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
        Type: XLSX.RelType.PivotCacheDefinition,
        Target: '../pivotCache/pivotCacheDefinition1.xml',
      },
    ]);
    zip.append(xml, {name: 'xl/pivotTables/_rels/pivotTable1.xml.rels'});
  }

  async addContentTypes(zip, model) {
    const xform = new ContentTypesXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: '[Content_Types].xml'});
  }

  async addApp(zip, model) {
    const xform = new AppXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: 'docProps/app.xml'});
  }

  async addCore(zip, model) {
    const coreXform = new CoreXform();
    zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
  }

  async addThemes(zip, model) {
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

  async addOfficeRels(zip) {
    const xform = new RelationshipsXform();
    const xml = xform.toXml([
      {Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml'},
      {Id: 'rId2', Type: XLSX.RelType.CoreProperties, Target: 'docProps/core.xml'},
      {Id: 'rId3', Type: XLSX.RelType.ExtenderProperties, Target: 'docProps/app.xml'},
    ]);
    zip.append(xml, {name: '_rels/.rels'});
  }

  async addWorkbookRels(zip, model) {
    let count = 1;
    const relationships = [
      {Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (model.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: XLSX.RelType.SharedStrings,
        Target: 'sharedStrings.xml',
      });
    }
    if ((model.pivotTables || []).length) {
      const pivotTable = model.pivotTables[0];
      pivotTable.rId = `rId${count++}`;
      relationships.push({
        Id: pivotTable.rId,
        Type: XLSX.RelType.PivotCacheDefinition,
        Target: 'pivotCache/pivotCacheDefinition1.xml',
      });
    }
    model.worksheets.forEach(worksheet => {
      worksheet.rId = `rId${count++}`;
      relationships.push({
        Id: worksheet.rId,
        Type: XLSX.RelType.Worksheet,
        Target: `worksheets/sheet${worksheet.id}.xml`,
      });
    });
    const xform = new RelationshipsXform();
    const xml = xform.toXml(relationships);
    zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
  }

  async addSharedStrings(zip, model) {
    if (model.sharedStrings && model.sharedStrings.count) {
      zip.append(model.sharedStrings.xml, {name: 'xl/sharedStrings.xml'});
    }
  }

  async addStyles(zip, model) {
    const {xml} = model.styles;
    if (xml) {
      zip.append(xml, {name: 'xl/styles.xml'});
    }
  }

  async addWorkbook(zip, model) {
    const xform = new WorkbookXform();
    zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
  }

  async addWorksheets(zip, model) {
    // preparation phase
    const worksheetXform = new WorksheetXform();
    const relationshipsXform = new RelationshipsXform();

    if ((model.worksheets || []).some(sheetNeedsCf)) {
      await worksheetXform.installCfXforms();
    }

    const hasComments = model.worksheets.some(ws => ws.comments && ws.comments.length > 0);
    let commentsXform;
    let vmlNotesXform;
    if (hasComments) {
      const CommentsXform = await loadCommentsXform();
      const VmlNotesXform = await loadVmlNotesXform();
      commentsXform = new CommentsXform();
      vmlNotesXform = new VmlNotesXform();
    }

    // write sheets
    model.worksheets.forEach(worksheet => {
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

  _finalize(zip) {
    return new Promise((resolve, reject) => {
      zip.on('finish', () => {
        resolve(this);
      });
      zip.on('error', reject);
      zip.finalize();
    });
  }

  async prepareModel(model, options) {
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
    model.styles = model.useStyles ? new StylesXform(true) : new StylesXform.Mock();

    // prepare all of the things before the render
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform();

    if ((model.worksheets || []).some(sheetNeedsCf)) {
      await worksheetXform.installCfXforms();
    }

    workbookXform.prepare(model);

    const worksheetOptions = {
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
    model.worksheets.forEach(worksheet => {
      // assign unique filenames to tables
      const tables = worksheet.tables
        ? Array.isArray(worksheet.tables)
          ? worksheet.tables
          : Object.values(worksheet.tables)
        : [];
      tables.forEach(table => {
        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
        model.tables.push(table);
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  }

  async write(stream, options) {
    options = options || {};
    const {model} = this.workbook;
    const zip = new ZipStream.ZipWriter(options.zip);
    zip.pipe(stream);

    await this.prepareModel(model, options);

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
    return this._finalize(zip);
  }

  writeFile(filename, options) {
    const stream = fs.createWriteStream(filename);

    return new Promise((resolve, reject) => {
      stream.on('finish', () => {
        resolve();
      });
      stream.on('error', error => {
        reject(error);
      });

      this.write(stream, options)
        .then(() => {
          stream.end();
        })
        .catch(err => {
          reject(err);
        });
    });
  }

  async writeBuffer(options) {
    const stream = new StreamBuf();
    await this.write(stream, options);
    return stream.read();
  }
}

XLSX.RelType = __esm_0;

export default XLSX;
export {XLSX};
