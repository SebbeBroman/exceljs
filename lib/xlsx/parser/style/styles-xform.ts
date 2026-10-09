import type {StylesIndex, StylesModel} from '../../xform/style/styles-xform.js';
export type {StylesIndex, StylesModel} from '../../xform/style/styles-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import StaticXform from '../static-xform.js';
import ListXform from '../list-xform.js';
import FontXform from './font-xform.js';
import FillXform from './fill-xform.js';
import BorderXform from './border-xform.js';
import NumFmtXform from './numfmt-xform.js';
import StyleXform from './style-xform.js';
import type {StyleModel} from './style-xform.js';
import DxfXform from './dxf-xform.js';
import type {DxfModel} from './dxf-xform.js';

// =============================================================================
// StylesXform is used to parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
class StylesXform extends BaseXform<StylesModel> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: Record<string, any>;
  index?: StylesIndex;

  constructor(initialise?: boolean) {
    super();

    this.map = {
      numFmts: new ListXform({tag: 'numFmts', count: true, childXform: new NumFmtXform()}),
      fonts: new ListXform({
        tag: 'fonts',
        count: true,
        childXform: new FontXform(),
        $: {'x14ac:knownFonts': 1},
      }),
      fills: new ListXform({tag: 'fills', count: true, childXform: new FillXform()}),
      borders: new ListXform({tag: 'borders', count: true, childXform: new BorderXform()}),
      cellStyleXfs: new ListXform({tag: 'cellStyleXfs', count: true, childXform: new StyleXform()}),
      cellXfs: new ListXform({
        tag: 'cellXfs',
        count: true,
        childXform: new StyleXform({xfId: true}),
      }),
      dxfs: new ListXform({tag: 'dxfs', always: true, count: true, childXform: new DxfXform()}),

      cellStyles: StylesXform.STATIC_XFORMS.cellStyles,
      tableStyles: StylesXform.STATIC_XFORMS.tableStyles,
      extLst: StylesXform.STATIC_XFORMS.extLst,
    };

    if (initialise) {
      // StylesXform also acts as style manager and is used to build up styles-model during worksheet processing
      this.init();
    }
  }

  initIndex(): void {
    this.index = {
      style: {},
      numFmt: {},
      numFmtNextId: 164, // start custom format ids here
      font: {},
      border: {},
      fill: {},
    };
  }

  init(): void {
    // Prepare for Style Manager role
    this.model = {
      styles: [],
      numFmts: [],
      fonts: [],
      borders: [],
      fills: [],
      dxfs: [],
    };

    this.initIndex();
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'styleSheet':
        this.initIndex();
        return true;
      default:
        this.parser = this.map[node.name] as BaseXform;
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
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
      case 'styleSheet': {
        this.model = {} as StylesModel;
        const add = (propName: keyof StylesModel, xform: BaseXform): void => {
          if (xform.model && (xform.model as unknown[]).length) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (this.model as any)[propName] = xform.model;
          }
        };
        add('numFmts', this.map.numFmts as BaseXform);
        add('fonts', this.map.fonts as BaseXform);
        add('fills', this.map.fills as BaseXform);
        add('borders', this.map.borders as BaseXform);
        add('styles', this.map.cellXfs as BaseXform);
        add('dxfs', this.map.dxfs as BaseXform);

        // index numFmts
        this.index = {
          model: [],
          numFmt: {},
        };
        if (this.model.numFmts) {
          const numFmtIndex = this.index.numFmt as Record<string | number, string | number>;
          this.model.numFmts.forEach(numFmt => {
            const nf = numFmt as {id: number; formatCode: string};
            numFmtIndex[nf.id] = nf.formatCode;
          });
        }

        return false;
      }
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(id: number): Record<string, unknown> | null {
    // if the style doesn't exist return null
    const style = this.model!.styles[id] as StyleModel | undefined;
    if (!style) return null;

    // have we built this model before?
    let model = this.index!.model![id];
    if (model) return model;

    // build a new model
    model = this.index!.model![id] = {};

    // -------------------------------------------------------
    // number format
    if (style.numFmtId) {
      const numFmt =
        (this.index!.numFmt as Record<string | number, string | number>)[style.numFmtId] ||
        NumFmtXform.getDefaultFmtCode(style.numFmtId);
      if (numFmt) {
        model.numFmt = numFmt;
      }
    }

    const addStyle = (
      name: string,
      group: unknown[] | undefined,
      styleId: number | undefined,
    ): void => {
      if (styleId || styleId === 0) {
        const part = group?.[styleId];
        if (part) {
          model![name] = part;
        }
      }
    };

    addStyle('font', this.model!.fonts, style.fontId);
    addStyle('border', this.model!.borders, style.borderId);
    addStyle('fill', this.model!.fills, style.fillId);

    // -------------------------------------------------------
    // alignment
    if (style.alignment) {
      model.alignment = style.alignment;
    }

    // -------------------------------------------------------
    // protection
    if (style.protection) {
      model.protection = style.protection;
    }

    return model;
  }

  getDxfStyle(id: number): DxfModel {
    return this.model!.dxfs[id];
  }

  // =========================================================================

  static STATIC_XFORMS = {
    cellStyles: new StaticXform({tag: 'cellStyles'}),
    tableStyles: new StaticXform({tag: 'tableStyles'}),
    extLst: new StaticXform({tag: 'extLst'}),
  };
}
export default StylesXform;
export {StylesXform};
