import {fromReadable, once} from '../utils/async-iterator.js';
import {entryToBuffer, entryToString, unzipToFiles} from '../utils/zip-reader.js';
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
import AppXform from './xform/core/app-xform.js';
import WorkbookXform from './xform/book/workbook-xform.js';
import WorksheetXform, {buildSheetRelMaps} from './xform/sheet/worksheet-xform.js';
import {
  loadDrawingXform,
  loadTableXform,
  loadCommentsXform,
  loadVmlNotesXform,
} from './lazy-xforms.js';

/** Workbook host that owns the model XLSX reads into / writes from */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type XlsxWorkbookHost = {model: any};

export interface XlsxReadOptions {
  ignoreNodes?: string[];
  base64?: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type XlsxModel = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type XmlSource = any;

function sheetNeedsCf(worksheet: XlsxModel): boolean {
  return Boolean(
    worksheet && worksheet.conditionalFormattings && worksheet.conditionalFormattings.length,
  );
}

// The fused cell reader is never needed when exporting a workbook.
// Cache the import promise so concurrent worksheet reads share its first load.
let fastSheetDataPromise: Promise<typeof import('./xform/sheet/fast-sheet-data.js')> | undefined;
function loadFastSheetData(): Promise<typeof import('./xform/sheet/fast-sheet-data.js')> {
  return (fastSheetDataPromise ??= import('./xform/sheet/fast-sheet-data.js'));
}

class XlsxReader {
  workbook: XlsxWorkbookHost;

  constructor(workbook: XlsxWorkbookHost) {
    this.workbook = workbook;
  }

  parseRels(stream: XmlSource): Promise<unknown> {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  }

  parseWorkbook(stream: XmlSource): Promise<XlsxModel> {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  }

  parseSharedStrings(stream: XmlSource): Promise<unknown> {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  }

  async reconcile(model: XlsxModel, options?: XlsxReadOptions): Promise<void> {
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform(options);

    // Load CF xforms if any sheet has conditional formatting (after parse they may)
    const needsCf = (model.worksheets || []).some(sheetNeedsCf);
    if (needsCf) {
      await worksheetXform.installCfXforms();
    }

    workbookXform.reconcile(model);

    // reconcile drawings with their rels (lazy-load drawing xform only if present)
    const drawingNames = Object.keys(model.drawings);
    if (drawingNames.length) {
      const DrawingXform = await loadDrawingXform();
      const drawingXform = new DrawingXform();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const drawingOptions: any = {
        media: model.media,
        mediaIndex: model.mediaIndex,
      };
      drawingNames.forEach(name => {
        const drawing = model.drawings[name];
        const drawingRel = model.drawingRels[name];
        if (drawingRel) {
          drawingOptions.rels = drawingRel.reduce(
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (o: any, rel: any) => {
              o[rel.Id] = rel;
              return o;
            },
            {},
          );
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (drawing.anchors || []).forEach((anchor: any) => {
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
    (model.worksheets as XlsxModel[]).forEach((worksheet: XlsxModel) => {
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

  /**
   * Fused fast path: sheet structure (cols, merges, validations, CF, page
   * setup, …) parses with the regular xforms over head/tail, while
   * `<sheetData>` cells parse + reconcile in one saxen pass. Falls back to
   * the classic path when the sheet is ineligible or fusion throws —
   * correctness first, speed second.
   */
  async _processWorksheetFast(
    split: {head: string; content: string; tail: string},
    xform: WorksheetXform,
    model: XlsxModel,
    sheetNo: string | number,
  ): Promise<XlsxModel | null> {
    const worksheet = (await xform.parseStream([
      `${split.head}<sheetData></sheetData>${split.tail}`,
    ])) as XlsxModel | null | undefined;
    if (!worksheet) {
      return null;
    }
    const {hyperlinkMap, commentsMap} = buildSheetRelMaps(
      worksheet,
      model.worksheetRels?.[sheetNo as string],
      model.comments,
      model.vmlDrawings,
    );
    const styles = model.styles as
      | {getStyleModel(id: number): Record<string, unknown> | null}
      | undefined;
    const fastCtx = {
      sharedStrings: model.sharedStrings,
      getStyleModel: (id: number) => styles?.getStyleModel(id),
      // Mirrors row reconcile: throws when styles are absent but used.
      getRowStyleModel: (id: number) => styles!.getStyleModel(id),
      date1904: model.properties && model.properties.date1904,
      hyperlinkMap,
      commentsMap,
      formulae: {},
    };
    // Fused saxen pass (parse + reconcile); classic path on bail/throw.
    const {parseFastSheetData} = await loadFastSheetData();
    const fastRows = parseFastSheetData(split.content, fastCtx);
    worksheet.rows = fastRows;
    worksheet.fastRows = true;
    return worksheet;
  }

  async _processWorksheetEntry(
    entryBytes: Uint8Array,
    model: XlsxModel,
    sheetNo: string | number,
    options: XlsxReadOptions | undefined,
    path: string,
  ): Promise<void> {
    const xform = new WorksheetXform(options);
    // Always install CF parsers for read — sheets may contain conditionalFormatting/extLst.
    // loadCfXforms is module-cached; concurrent sheets share one import.
    await xform.installCfXforms();
    const xml = entryToString(entryBytes);

    const {isFastSheetDataEnabled, splitSheetData, canUseFastSheetData} = await loadFastSheetData();
    let worksheet: XlsxModel | null | undefined;
    const fastEligible =
      isFastSheetDataEnabled() &&
      !options?.ignoreNodes?.length &&
      !(options as {maxRows?: number; maxCols?: number} | undefined)?.maxRows &&
      !(options as {maxRows?: number; maxCols?: number} | undefined)?.maxCols;
    if (fastEligible) {
      const split = splitSheetData(xml);
      if (split && canUseFastSheetData(split.content)) {
        try {
          worksheet = await this._processWorksheetFast(split, xform, model, sheetNo);
        } catch {
          worksheet = undefined;
          // Fall through to the classic path below.
        }
      }
    }
    if (!worksheet) {
      worksheet = (await xform.parseStream([xml])) as XlsxModel | null | undefined;
    }
    if (!worksheet) {
      return;
    }
    worksheet.sheetNo = sheetNo;
    model.worksheetHash[path] = worksheet;
    // reconcile() rebuilds worksheets from worksheetHash in workbook sheet order,
    // so concurrent push order does not matter for correctness.
    model.worksheets.push(worksheet);
  }

  async _processCommentEntry(stream: XmlSource, model: XlsxModel, name: string): Promise<void> {
    const CommentsXform = await loadCommentsXform();
    const xform = new CommentsXform();
    const comments = await xform.parseStream(stream);
    model.comments[`../${name}.xml`] = comments;
  }

  async _processTableEntry(stream: XmlSource, model: XlsxModel, name: string): Promise<void> {
    const TableXform = await loadTableXform();
    const xform = new TableXform();
    const table = await xform.parseStream(stream);
    model.tables[`../tables/${name}.xml`] = table;
  }

  async _processWorksheetRelsEntry(
    stream: XmlSource,
    model: XlsxModel,
    sheetNo: string | number,
  ): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.worksheetRels[sheetNo] = relationships;
  }

  async _processMediaEntry(entry: XmlSource, model: XlsxModel, filename: string): Promise<void> {
    const lastDot = filename.lastIndexOf('.');
    // if we can't determine extension, ignore it
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      const chunks: Uint8Array[] = [];
      for await (const chunk of fromReadable(entry)) {
        chunks.push(
          isBytes(chunk)
            ? chunk
            : bytesFrom(chunk as string | ArrayBuffer | ArrayLike<number> | ArrayBufferView),
        );
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

  async _processDrawingEntry(entry: XmlSource, model: XlsxModel, name: string): Promise<void> {
    const DrawingXform = await loadDrawingXform();
    const xform = new DrawingXform();
    const drawing = await xform.parseStream(entry);
    model.drawings[name] = drawing;
  }

  async _processDrawingRelsEntry(entry: XmlSource, model: XlsxModel, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.drawingRels[name] = relationships;
  }

  async _processVmlDrawingEntry(entry: XmlSource, model: XlsxModel, name: string): Promise<void> {
    const VmlNotesXform = await loadVmlNotesXform();
    const xform = new VmlNotesXform();
    const vmlDrawing = await xform.parseStream(entry);
    model.vmlDrawings[`../drawings/${name}.vml`] = vmlDrawing;
  }

  async _processThemeEntry(entry: XmlSource, model: XlsxModel, name: string): Promise<void> {
    // Theme is stored as raw XML text for round-trip rewrite fidelity.
    // Prefer a single string chunk (buffered load); fall back to concat for streams.
    if (typeof entry === 'string') {
      model.themes[name] = entry;
      return;
    }
    const chunks: Uint8Array[] = [];
    let sawString = false;
    let stringAcc = '';
    for await (const chunk of fromReadable(entry)) {
      if (typeof chunk === 'string') {
        sawString = true;
        stringAcc += chunk;
      } else {
        chunks.push(
          isBytes(chunk)
            ? chunk
            : bytesFrom(chunk as string | ArrayBuffer | ArrayLike<number> | ArrayBufferView),
        );
      }
    }
    if (sawString && chunks.length === 0) {
      model.themes[name] = stringAcc;
    } else if (chunks.length) {
      model.themes[name] = bytesToString(concat(chunks));
    } else {
      model.themes[name] = stringAcc;
    }
  }

  async load(
    data: Uint8Array | ArrayBuffer | ArrayBufferView | string | Buffer,
    options?: XlsxReadOptions,
  ): Promise<XlsxWorkbookHost> {
    let buffer: Uint8Array | Buffer | string;
    if (options && options.base64) {
      buffer = bytesFrom(String(data), 'base64');
    } else if (isBytes(data)) {
      buffer = data;
    } else if (data instanceof ArrayBuffer) {
      buffer = asUint8Array(data);
    } else if (ArrayBuffer.isView(data)) {
      buffer = asUint8Array(data);
    } else {
      buffer = data as unknown as Uint8Array;
    }

    const model: XlsxModel = {
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

    const zipFiles = await unzipToFiles(buffer as ArrayBuffer | ArrayBufferView);

    // Two phases so sheets never wait on each other and never wait on the
    // global reconcile barrier for their dependencies:
    //  1. package-level parts (workbook, rels, shared strings, styles, media,
    //     comments, …) — all small; parsed concurrently.
    //  2. sheets — each parses (fused fast path when eligible) with SST,
    //     styles, rels and comments already resolved, flowing straight into
    //     reconcile-ready models. Sheets pipeline past each other.
    const partTasks: Promise<void>[] = [];
    const sheetEntries: Array<{entryBytes: Uint8Array; sheetNo: string; path: string}> = [];

    for (const rawName of Object.keys(zipFiles)) {
      // fflate omits pure directory entries; still skip trailing-slash keys if present
      if (!rawName || rawName.endsWith('/')) {
        continue;
      }
      const entryName = rawName[0] === '/' ? rawName.slice(1) : rawName;
      const entryBytes = zipFiles[rawName]!;

      const sheetMatch = /^xl\/worksheets\/sheet(\d+)[.]xml$/.exec(entryName);
      if (sheetMatch) {
        sheetEntries.push({entryBytes, sheetNo: sheetMatch[1]!, path: entryName});
        continue;
      }

      // Media keeps binary; theme is kept as raw XML text for rewrite fidelity.
      const isMedia = /xl\/media\//.test(entryName);
      const isTheme = /xl\/theme\/[a-zA-Z0-9]+[.]xml/.test(entryName);

      partTasks.push(
        (async () => {
          let xmlSource: XmlSource;
          if (isMedia) {
            xmlSource = once(entryToBuffer(entryBytes));
          } else if (isTheme) {
            // Decode once to string — skip buffer→chunk→concat round-trip.
            xmlSource = entryToString(entryBytes);
          } else {
            // Sync one-shot iterable: full string is already in memory after unzip.
            // Avoids async-generator overhead from stringChunks for the buffered path.
            xmlSource = [entryToString(entryBytes)];
          }

          switch (entryName) {
            case '_rels/.rels':
              model.globalRels = await this.parseRels(xmlSource);
              return;

            case 'xl/workbook.xml': {
              const workbook = await this.parseWorkbook(xmlSource);
              model.sheets = workbook.sheets;
              model.definedNames = workbook.definedNames;
              model.views = workbook.views;
              model.properties = workbook.properties;
              model.calcProperties = workbook.calcProperties;
              return;
            }

            case 'xl/_rels/workbook.xml.rels':
              model.workbookRels = await this.parseRels(xmlSource);
              return;

            case 'xl/sharedStrings.xml':
              model.sharedStrings = new SharedStringsXform();
              await model.sharedStrings.parseStream(xmlSource);
              return;

            case 'xl/styles.xml':
              model.styles = new StylesXform();
              await model.styles.parseStream(xmlSource);
              return;

            case 'docProps/app.xml': {
              const appXform = new AppXform();
              const appProperties = await appXform.parseStream(xmlSource);
              model.company = appProperties?.company;
              model.manager = appProperties?.manager;
              return;
            }

            case 'docProps/core.xml': {
              const coreXform = new CoreXform();
              const coreProperties = await coreXform.parseStream(xmlSource);
              Object.assign(model, coreProperties);
              return;
            }

            default: {
              let match = entryName.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
              if (match) {
                await this._processWorksheetRelsEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
              if (match) {
                await this._processThemeEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
              if (match) {
                await this._processMediaEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
              if (match) {
                await this._processDrawingEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/(comments\d+)[.]xml/);
              if (match) {
                await this._processCommentEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/tables\/(table\d+)[.]xml/);
              if (match) {
                await this._processTableEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
              if (match) {
                await this._processDrawingRelsEntry(xmlSource, model, match[1]);
                return;
              }
              match = entryName.match(/xl\/drawings\/(vmlDrawing\d+)[.]vml/);
              if (match) {
                await this._processVmlDrawingEntry(xmlSource, model, match[1]);
                return;
              }
            }
          }
        })(),
      );
    }

    await Promise.all(partTasks);

    // Phase 2: sheets pipeline past each other, each with dependencies ready.
    await Promise.all(
      sheetEntries.map(({entryBytes, sheetNo, path}) =>
        this._processWorksheetEntry(entryBytes, model, sheetNo, options, path),
      ),
    );

    // Media entries complete in nondeterministic order under Promise.all, but
    // each `_processMediaEntry` assigns index+push atomically so index and array
    // stay consistent. Re-sort by filename here for deterministic output bytes:
    // drawings resolve via mediaIndex (by name), so reordering + rebuilding the
    // index before reconcile() is safe.
    if (model.media.length > 1) {
      model.media.sort((a: XlsxModel, b: XlsxModel) =>
        `${a.name}.${a.extension}`.localeCompare(`${b.name}.${b.extension}`),
      );
      model.mediaIndex = {};
      model.media.forEach((medium: XlsxModel, i: number) => {
        model.mediaIndex[`${medium.name}.${medium.extension}`] = i;
        model.mediaIndex[medium.name] = i;
      });
    }

    await this.reconcile(model, options);

    // apply model
    this.workbook.model = model;
    return this.workbook;
  }
}

export default XlsxReader;
