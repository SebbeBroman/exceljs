import type {
  WorksheetMap,
  SheetRelMaps,
  WorksheetXformOptions,
  WorksheetXformConstructorOptions,
  WorksheetXformModel,
  WorksheetRel,
} from '../../xform/sheet/worksheet-xform.js';
export type {
  WorksheetMap,
  SheetRelMaps,
  WorksheetXformOptions,
  WorksheetXformConstructorOptions,
  WorksheetXformModel,
  WorksheetRel,
} from '../../xform/sheet/worksheet-xform.js';
import RelType from '../../rel-type.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
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
import {loadCfXforms} from '../../lazy-parsers.js';

const mergeRule = (rule: Record<string, unknown>, extRule: Record<string, unknown>): void => {
  Object.keys(extRule).forEach(key => {
    const value = rule[key];
    const extValue = extRule[key];
    if (value === undefined && extValue !== undefined) {
      rule[key] = extValue;
    }
  });
};

const mergeConditionalFormattings = (
  model: Array<Record<string, unknown>> | null | undefined,
  extModel: Array<Record<string, unknown>> | null | undefined,
): Array<Record<string, unknown>> | null | undefined => {
  // conditional formattings are rendered in worksheet.conditionalFormatting and also in
  // worksheet.extLst.ext.x14:conditionalFormattings
  // some (e.g. dataBar) are even spread across both!
  if (!extModel || !extModel.length) {
    return model;
  }
  if (!model || !model.length) {
    return extModel;
  }

  // index model rules by x14Id
  const cfMap: Record<string, Record<string, unknown>> = {};
  const ruleMap: Record<string, Record<string, unknown>> = {};
  model.forEach(cf => {
    cfMap[cf.ref as string] = cf;
    (cf.rules as Array<Record<string, unknown>>).forEach(rule => {
      const {x14Id} = rule;
      if (x14Id) {
        ruleMap[x14Id as string] = rule;
      }
    });
  });

  extModel.forEach(extCf => {
    (extCf.rules as Array<Record<string, unknown>>).forEach(extRule => {
      const rule = ruleMap[extRule.x14Id as string];
      if (rule) {
        // merge with matching rule
        mergeRule(rule, extRule);
      } else if (cfMap[extCf.ref as string]) {
        // reuse existing cf ref
        (cfMap[extCf.ref as string].rules as Array<Record<string, unknown>>).push(extRule);
      } else {
        // create new cf
        model.push({
          ref: extCf.ref,
          rules: [extRule],
        });
      }
    });
  });

  // need to cope with rules in extModel that don't exist in model
  return model;
};

/**
 * Build relationship / hyperlink / comment lookup maps for a parsed worksheet
 * model. Extracted from reconcile so the fast sheetData path can resolve
 * hyperlinks/comments inline; reconcile itself calls this (same behavior).
 */
export function buildSheetRelMaps(
  model: WorksheetXformModel,
  relationships: WorksheetRel[] | undefined,
  commentsByTarget: Record<string, {comments: Array<Record<string, unknown>>}> | undefined,
  vmlDrawings: Record<string, {comments: unknown[]}> | undefined,
): SheetRelMaps {
  const rels = (relationships || []).reduce(
    (h: Record<string, WorksheetRel>, rel: WorksheetRel) => {
      h[rel.Id] = rel;
      if (rel.Type === RelType.Comments) {
        model.comments = (commentsByTarget as unknown as Record<string, {comments: unknown[]}>)[
          rel.Target
        ].comments as Array<Record<string, unknown>>;
      }
      if (rel.Type === RelType.VmlDrawing && model.comments && model.comments.length) {
        const vmlComment = vmlDrawings![rel.Target].comments;
        model.comments.forEach((comment, index) => {
          comment.note = Object.assign({}, comment.note, vmlComment[index]);
        });
      }
      return h;
    },
    {},
  );
  const commentsMap = (model.comments || []).reduce(
    (h: Record<string, unknown>, comment: Record<string, unknown>) => {
      if (comment.ref) {
        h[comment.ref as string] = comment;
      }
      return h;
    },
    {},
  );
  const hyperlinkMap = (model.hyperlinks || []).reduce(
    (h: Record<string, string>, hyperlink: Record<string, unknown>) => {
      if (hyperlink.rId) {
        h[hyperlink.address as string] = rels[hyperlink.rId as string].Target;
      }
      return h;
    },
    {},
  );
  return {rels, hyperlinkMap, commentsMap};
}
class EmptyConditionalFormattingsXform extends BaseXform<unknown[]> {
  constructor() {
    super();
    this.model = [];
  }
  override reconcile(): void {}
  override reset(): void {
    this.model = [];
  }
  override parseOpen(): boolean {
    return false;
  }
  override parseText(): void {}
  override parseClose(): boolean {
    return false;
  }
}

class EmptyExtLstXform extends BaseXform<Record<string, unknown> | null> {
  constructor() {
    super();
    this.model = null;
  }
  hasContent(): boolean {
    return false;
  }
  override reset(): void {
    this.model = null;
  }
  override parseOpen(): boolean {
    return false;
  }
  override parseText(): void {}
  override parseClose(): boolean {
    return false;
  }
}

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
   * Install real CF + extLst xforms (dynamic import). Call before parse, or before
   * prepare/render when the sheet model has conditional formatting.
   */
  async installCfXforms(): Promise<void> {
    if (this._cfInstalled) return;
    const {ConditionalFormattingsXform, ExtLstXform} = await loadCfXforms();
    this.map.conditionalFormatting = new ConditionalFormattingsXform();
    this.map.extLst = new ExtLstXform();
    this._cfInstalled = true;
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (node.name === 'worksheet') {
      for (const xform of Object.values(this.map) as Array<
        BaseXform | EmptyConditionalFormattingsXform | EmptyExtLstXform
      >) {
        xform.reset();
      }
      return true;
    }

    if (this.map[node.name] && !this.ignoreNodes.includes(node.name)) {
      this.parser = this.map[node.name] as BaseXform;
      this.parser.parseOpen(node);
    }
    return true;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'worksheet': {
        const properties: Record<string, unknown> = {
          ...(this.map.sheetFormatPr.model as object),
        };
        const sheetPrModel = this.map.sheetPr.model as Record<string, unknown> | null | undefined;
        if (sheetPrModel && sheetPrModel.tabColor) {
          properties.tabColor = sheetPrModel.tabColor;
        }
        if (sheetPrModel && sheetPrModel.outlineProperties) {
          properties.outlineProperties = sheetPrModel.outlineProperties;
        }
        const sheetProperties = {
          fitToPage:
            (sheetPrModel &&
              (sheetPrModel.pageSetup as {fitToPage?: boolean} | undefined)?.fitToPage) ||
            false,
          margins: this.map.pageMargins.model,
        };
        const pageSetup = Object.assign(
          sheetProperties,
          this.map.pageSetup.model,
          this.map.printOptions.model,
        );
        const extLstModel = this.map.extLst.model as Record<string, unknown> | null | undefined;
        const conditionalFormattings = mergeConditionalFormattings(
          this.map.conditionalFormatting.model as Array<Record<string, unknown>>,
          extLstModel &&
            (extLstModel['x14:conditionalFormattings'] as Array<Record<string, unknown>>),
        );
        this.model = {
          dimensions: this.map.dimension.model as string | undefined,
          cols: this.map.cols.model,
          rows: this.map.sheetData.model as unknown[],
          mergeCells: this.map.mergeCells.model as string[] | undefined,
          hyperlinks: this.map.hyperlinks.model as Array<Record<string, unknown>>,
          dataValidations: this.map.dataValidations.model,
          properties,
          views: this.map.sheetViews.model,
          pageSetup,
          headerFooter: this.map.headerFooter.model,
          background: this.map.picture.model as {rId?: string} | undefined,
          drawing: this.map.drawing.model as Record<string, unknown> | undefined,
          tables: this.map.tableParts.model as Array<Record<string, unknown>>,
          conditionalFormattings: conditionalFormattings as unknown[],
        };

        if (this.map.autoFilter.model) {
          this.model.autoFilter = this.map.autoFilter.model;
        }
        if (this.map.sheetProtection.model) {
          this.model.sheetProtection = this.map.sheetProtection.model;
        }

        return false;
      }

      default:
        // not quite sure how we get here!
        return true;
    }
  }

  override reconcile(model?: WorksheetXformModel | null, options?: WorksheetXformOptions): void {
    if (!model || !options) {
      return;
    }
    const {rels, hyperlinkMap, commentsMap} = buildSheetRelMaps(
      model,
      model.relationships,
      options.comments as unknown as Record<string, {comments: Array<Record<string, unknown>>}>,
      options.vmlDrawings as unknown as Record<string, {comments: unknown[]}>,
    );
    options.commentsMap = commentsMap;
    options.hyperlinkMap = hyperlinkMap;
    options.formulae = {};

    // compact the rows and cells
    model.rows = (model.rows && (model.rows as unknown[]).filter(Boolean)) || [];
    (model.rows as Array<{cells?: unknown[]}>).forEach(row => {
      row.cells = (row.cells && row.cells.filter(Boolean)) || [];
    });

    this.map.cols.reconcile(model.cols as never, options);
    // Fast-path rows are already reconciled (see fast-sheet-data) — skip the
    // per-cell pass but still reconcile cols/CF below.
    if (!(model as {fastRows?: boolean}).fastRows) {
      this.map.sheetData.reconcile(model.rows as never, options);
    }
    this.map.conditionalFormatting.reconcile(model.conditionalFormattings as never, options);

    model.media = [];
    if (model.drawing) {
      const drawingRel = rels[model.drawing.rId as string];
      const match = drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
      if (match) {
        const drawingName = match[1];
        const drawing = (
          options.drawings as unknown as Record<string, {anchors: Array<Record<string, unknown>>}>
        )[drawingName];
        drawing.anchors.forEach(anchor => {
          if (anchor.medium) {
            const image = {
              type: 'image',
              imageId: (anchor.medium as {index: number}).index,
              range: anchor.range,
              hyperlinks: (anchor.picture as {hyperlinks?: unknown}).hyperlinks,
            };
            model.media!.push(image);
          }
        });
      }
    }

    const backgroundRel = model.background && rels[model.background.rId as string];
    if (backgroundRel) {
      const target = backgroundRel.Target.split('/media/')[1];
      const imageId = options.mediaIndex && options.mediaIndex[target];
      if (imageId !== undefined) {
        model.media!.push({
          type: 'background',
          imageId,
        });
      }
    }

    model.tables = (model.tables || []).map(tablePart => {
      const rel = rels[tablePart.rId as string];
      return options.tables![rel.Target] as Record<string, unknown>;
    });

    delete model.relationships;
    delete model.hyperlinks;
    delete model.comments;
  }
}

export default WorkSheetXform;
export {WorkSheetXform};
