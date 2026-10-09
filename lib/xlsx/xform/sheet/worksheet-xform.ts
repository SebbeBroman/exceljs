import type XformState from '../xform-state.js';
import colCache from '../../../utils/col-cache.js';
import XmlStream from '../../../utils/xml-stream.js';
import RelType from '../../rel-type.js';
import Merges from './merges.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import ListXform from '../list-xform.js';
import RowXform from './row-xform.js';
import ColXform from './col-xform.js';
import DimensionXform from './dimension-xform.js';
import HyperlinkXform from './hyperlink-xform.js';
import MergeCellXform from './merge-cell-xform.js';
import DataValidationsXform from './data-validations-xform.js';
import SheetPropertiesXform from './sheet-properties-xform.js';
import SheetFormatPropertiesXform from './sheet-format-properties-xform.js';
import SheetViewXform from './sheet-view-xform.js';
import SheetProtectionXform from './sheet-protection-xform.js';
import PageMarginsXform from './page-margins-xform.js';
import PageSetupXform from './page-setup-xform.js';
import PrintOptionsXform from './print-options-xform.js';
import AutoFilterXform from './auto-filter-xform.js';
import PictureXform from './picture-xform.js';
import DrawingXform from './drawing-xform.js';
import TablePartXform from './table-part-xform.js';
import RowBreaksXform from './row-breaks-xform.js';
import HeaderFooterXform from './header-footer-xform.js';
import {loadCfXforms} from '../../lazy-xforms.js';

export interface WorksheetRel {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

export interface WorksheetXformModel {
  id?: number;
  cols?: unknown;
  rows?: unknown[];
  dimensions?: string;
  views?: unknown;
  properties?: Record<string, unknown>;
  pageSetup?: Record<string, unknown>;
  headerFooter?: unknown;
  autoFilter?: unknown;
  mergeCells?: string[];
  hyperlinks?: Array<Record<string, unknown>>;
  comments?: Array<Record<string, unknown>>;
  dataValidations?: unknown;
  conditionalFormattings?: unknown[];
  sheetProtection?: unknown;
  rowBreaks?: unknown[];
  drawing?: Record<string, unknown>;
  background?: {rId?: string};
  image?: unknown;
  tables?: Array<Record<string, unknown>>;
  media?: Array<Record<string, unknown>>;
  pivotTables?: unknown[];
  rels?: WorksheetRel[];
  relationships?: WorksheetRel[];
  [key: string]: unknown;
}

export interface WorksheetXformConstructorOptions {
  maxRows?: number;
  maxCols?: number;
  ignoreNodes?: string[];
}

export interface WorksheetXformOptions extends XformOptions {
  merges?: Merges;
  hyperlinks?: Array<Record<string, unknown>>;
  comments?: Array<Record<string, unknown>>;
  formulae?: Record<string | number, unknown>;
  siFormulae?: number;
  commentRefs?: Array<{commentName: string; vmlDrawing: string}>;
  media?: Record<string | number, {name: string; extension: string; [key: string]: unknown}>;
  mediaIndex?: Record<string, number>;
  drawings?: Array<Record<string, unknown>>;
  drawingsCount?: number;
  styles?: {
    addDxfStyle(style: unknown): number;
  };
  commentsMap?: Record<string, unknown>;
  hyperlinkMap?: Record<string, string>;
  tables?: Record<string, unknown>;
  vmlDrawings?: Record<string, {comments: unknown[]}>;
  [key: string]: unknown;
}

export interface SheetRelMaps {
  rels: Record<string, WorksheetRel>;
  hyperlinkMap: Record<string, string>;
  commentsMap: Record<string, unknown>;
}
class EmptyConditionalFormattingsXform extends BaseXform<unknown[]> {
  constructor() {
    super();
    this.model = [];
  }

  override prepare(): void {}
  override render(): void {}
  override reset(): void {
    this.model = [];
  }
}

class EmptyExtLstXform extends BaseXform<Record<string, unknown> | null> {
  constructor() {
    super();
    this.model = null;
  }

  override prepare(): void {}
  override render(): void {}
  hasContent(): boolean {
    return false;
  }
  override reset(): void {
    this.model = null;
  }
}

export type WorksheetMap = Record<string, XformState>;

class WorkSheetXform extends BaseXform<WorksheetXformModel> {
  ignoreNodes: string[];
  _cfInstalled: boolean;
  declare map: WorksheetMap;
  preImageId?: number | string;

  constructor(options?: WorksheetXformConstructorOptions) {
    super();

    const {maxRows, maxCols, ignoreNodes} = options || {};

    this.ignoreNodes = ignoreNodes || [];
    this._cfInstalled = false;

    this.map = {
      sheetPr: new SheetPropertiesXform(),
      dimension: new DimensionXform(),
      sheetViews: new ListXform({
        tag: 'sheetViews',
        count: false,
        childXform: new SheetViewXform(),
      }),
      sheetFormatPr: new SheetFormatPropertiesXform(),
      cols: new ListXform({tag: 'cols', count: false, childXform: new ColXform()}),
      sheetData: new ListXform({
        tag: 'sheetData',
        count: false,
        empty: true,
        childXform: new RowXform({maxItems: maxCols}),
        maxItems: maxRows,
      }),
      autoFilter: new AutoFilterXform(),
      mergeCells: new ListXform({tag: 'mergeCells', count: true, childXform: new MergeCellXform()}),
      rowBreaks: new RowBreaksXform(),
      hyperlinks: new ListXform({
        tag: 'hyperlinks',
        count: false,
        childXform: new HyperlinkXform(),
      }),
      pageMargins: new PageMarginsXform(),
      dataValidations: new DataValidationsXform(),
      pageSetup: new PageSetupXform(),
      headerFooter: new HeaderFooterXform(),
      printOptions: new PrintOptionsXform(),
      picture: new PictureXform(),
      drawing: new DrawingXform(),
      sheetProtection: new SheetProtectionXform(),
      tableParts: new ListXform({tag: 'tableParts', count: true, childXform: new TablePartXform()}),
      // Stubs until installCfXforms() — keeps CF out of write-only bundles
      conditionalFormatting: new EmptyConditionalFormattingsXform(),
      extLst: new EmptyExtLstXform(),
    };
  }

  /**
   * Install real CF + extLst xforms (dynamic import). Call before
   * prepare/render when the sheet model has conditional formatting.
   */
  async installCfXforms(): Promise<void> {
    if (this._cfInstalled) return;
    const {ConditionalFormattingsXform, ExtLstXform} = await loadCfXforms();
    this.map.conditionalFormatting = new ConditionalFormattingsXform();
    this.map.extLst = new ExtLstXform();
    this._cfInstalled = true;
  }

  override prepare(model?: WorksheetXformModel | null, options?: WorksheetXformOptions): void {
    if (!model || !options) {
      return;
    }
    options.merges = new Merges();
    model.hyperlinks = options.hyperlinks = [];
    model.comments = options.comments = [];

    options.formulae = {};
    options.siFormulae = 0;
    this.map.cols.prepare(model.cols as never, options);
    this.map.sheetData.prepare(model.rows as never, options);
    this.map.conditionalFormatting.prepare(model.conditionalFormattings as never, options);

    model.mergeCells = options.merges.mergeCells;

    // prepare relationships
    const rels = (model.rels = [] as WorksheetRel[]);

    function nextRid(r: WorksheetRel[]): string {
      return `rId${r.length + 1}`;
    }

    model.hyperlinks.forEach(hyperlink => {
      const rId = nextRid(rels);
      hyperlink.rId = rId;
      rels.push({
        Id: rId,
        Type: RelType.Hyperlink,
        Target: hyperlink.target as string,
        TargetMode: 'External',
      });
    });

    // prepare comment relationships
    if (model.comments.length > 0) {
      const comment = {
        Id: nextRid(rels),
        Type: RelType.Comments,
        Target: `../comments${model.id}.xml`,
      };
      rels.push(comment);
      const vmlDrawing = {
        Id: nextRid(rels),
        Type: RelType.VmlDrawing,
        Target: `../drawings/vmlDrawing${model.id}.vml`,
      };
      rels.push(vmlDrawing);

      model.comments.forEach(item => {
        item.refAddress = colCache.decodeAddress(item.ref as string);
      });

      options.commentRefs!.push({
        commentName: `comments${model.id}`,
        vmlDrawing: `vmlDrawing${model.id}`,
      });
    }

    const drawingRelsHash: string[] = [];
    let bookImage: {name: string; extension: string; [key: string]: unknown} | undefined;
    (model.media || []).forEach(medium => {
      if (medium.type === 'background') {
        const rId = nextRid(rels);
        bookImage = options.media![medium.imageId as number];
        rels.push({
          Id: rId,
          Type: RelType.Image,
          Target: `../media/${bookImage.name}.${bookImage.extension}`,
        });
        model.background = {
          rId,
        };
        model.image = options.media![medium.imageId as number];
      } else if (medium.type === 'image') {
        let {drawing} = model;
        bookImage = options.media![medium.imageId as number];
        if (!drawing) {
          drawing = model.drawing = {
            rId: nextRid(rels),
            name: `drawing${++(options.drawingsCount as number)}`,
            anchors: [],
            rels: [],
          };
          options.drawings!.push(drawing);
          rels.push({
            Id: drawing.rId as string,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
            Target: `../drawings/${drawing.name}.xml`,
          });
        }
        let rIdImage =
          this.preImageId === medium.imageId
            ? drawingRelsHash[medium.imageId as number]
            : drawingRelsHash[(drawing.rels as WorksheetRel[]).length];
        if (!rIdImage) {
          rIdImage = nextRid(drawing.rels as WorksheetRel[]);
          drawingRelsHash[(drawing.rels as WorksheetRel[]).length] = rIdImage;
          (drawing.rels as WorksheetRel[]).push({
            Id: rIdImage,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            Target: `../media/${bookImage.name}.${bookImage.extension}`,
          });
        }

        const anchor: Record<string, unknown> = {
          picture: {
            rId: rIdImage,
          },
          range: medium.range,
        };
        if (medium.hyperlinks && (medium.hyperlinks as {hyperlink?: string}).hyperlink) {
          const rIdHyperLink = nextRid(drawing.rels as WorksheetRel[]);
          drawingRelsHash[(drawing.rels as WorksheetRel[]).length] = rIdHyperLink;
          (anchor.picture as Record<string, unknown>).hyperlinks = {
            tooltip: (medium.hyperlinks as {tooltip?: string}).tooltip,
            rId: rIdHyperLink,
          };
          (drawing.rels as WorksheetRel[]).push({
            Id: rIdHyperLink,
            Type: RelType.Hyperlink,
            Target: (medium.hyperlinks as {hyperlink: string}).hyperlink,
            TargetMode: 'External',
          });
        }
        this.preImageId = medium.imageId as number | string;
        (drawing.anchors as unknown[]).push(anchor);
      }
    });

    // prepare tables
    (model.tables || []).forEach(table => {
      // relationships
      const rId = nextRid(rels);
      table.rId = rId;
      rels.push({
        Id: rId,
        Type: RelType.Table,
        Target: `../tables/${table.target}`,
      });

      // dynamic styles
      (table.columns as Array<Record<string, unknown>>).forEach(column => {
        const {style} = column;
        if (style) {
          column.dxfId = options.styles!.addDxfStyle(style);
        }
      });
    });

    // prepare pivot tables
    if ((model.pivotTables || []).length) {
      rels.push({
        Id: nextRid(rels),
        Type: RelType.PivotTable,
        Target: '../pivotTables/pivotTable1.xml',
      });
    }

    // prepare ext items
    this.map.extLst.prepare(model as never, options);
  }

  override render(xmlStream: XmlStreamLike, model?: WorksheetXformModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openXml(
      (XmlStream as unknown as {StdDocAttributes: Record<string, string>}).StdDocAttributes,
    );
    xmlStream.openNode('worksheet', WorkSheetXform.WORKSHEET_ATTRIBUTES);

    const sheetFormatPropertiesModel: Record<string, unknown> | undefined = model.properties
      ? {
          defaultRowHeight: model.properties.defaultRowHeight,
          dyDescent: model.properties.dyDescent,
          outlineLevelCol: model.properties.outlineLevelCol,
          outlineLevelRow: model.properties.outlineLevelRow,
        }
      : undefined;
    if (model.properties && model.properties.defaultColWidth) {
      sheetFormatPropertiesModel!.defaultColWidth = model.properties.defaultColWidth;
    }
    const sheetPropertiesModel = {
      outlineProperties: model.properties && model.properties.outlineProperties,
      tabColor: model.properties && model.properties.tabColor,
      pageSetup:
        model.pageSetup && model.pageSetup.fitToPage
          ? {
              fitToPage: model.pageSetup.fitToPage,
            }
          : undefined,
    };
    const pageMarginsModel = model.pageSetup && model.pageSetup.margins;
    const printOptionsModel = {
      showRowColHeaders: model.pageSetup && model.pageSetup.showRowColHeaders,
      showGridLines: model.pageSetup && model.pageSetup.showGridLines,
      horizontalCentered: model.pageSetup && model.pageSetup.horizontalCentered,
      verticalCentered: model.pageSetup && model.pageSetup.verticalCentered,
    };
    const sheetProtectionModel = model.sheetProtection;

    this.map.sheetPr.render(xmlStream, sheetPropertiesModel as never);
    this.map.dimension.render(xmlStream, model.dimensions as never);
    this.map.sheetViews.render(xmlStream, model.views as never);
    this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel as never);
    this.map.cols.render(xmlStream, model.cols as never);
    this.map.sheetData.render(xmlStream, model.rows as never);
    this.map.sheetProtection.render(xmlStream, sheetProtectionModel as never); // Note: must be after sheetData and before autoFilter
    this.map.autoFilter.render(xmlStream, model.autoFilter as never);
    this.map.mergeCells.render(xmlStream, model.mergeCells as never);
    this.map.conditionalFormatting.render(xmlStream, model.conditionalFormattings as never); // Note: must be before dataValidations
    this.map.dataValidations.render(xmlStream, model.dataValidations as never);

    // For some reason hyperlinks have to be after the data validations
    this.map.hyperlinks.render(xmlStream, model.hyperlinks as never);

    this.map.printOptions.render(xmlStream, printOptionsModel as never); // Note: must be before pageMargins
    this.map.pageMargins.render(xmlStream, pageMarginsModel as never);
    this.map.pageSetup.render(xmlStream, model.pageSetup as never);
    this.map.headerFooter.render(xmlStream, model.headerFooter as never);
    this.map.rowBreaks.render(xmlStream, model.rowBreaks as never);
    this.map.drawing.render(xmlStream, model.drawing as never); // Note: must be after rowBreaks
    this.map.picture.render(xmlStream, model.background as never); // Note: must be after drawing
    this.map.tableParts.render(xmlStream, model.tables as never);

    this.map.extLst.render(xmlStream, model as never);

    if (model.rels) {
      // add a <legacyDrawing /> node for each comment
      model.rels.forEach(rel => {
        if (rel.Type === RelType.VmlDrawing) {
          xmlStream.leafNode('legacyDrawing', {'r:id': rel.Id});
        }
      });
    }

    xmlStream.closeNode();
  }

  static WORKSHEET_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x14ac',
    'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
  };
}

export default WorkSheetXform;
export {WorkSheetXform};
