import Enums from '../../../model/enums.js';
import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
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

// custom numfmt ids start here
const NUMFMT_BASE = 164;

export interface StylesModel {
  styles: (string | StyleModel)[];
  numFmts: (string | {id: number; formatCode: string})[];
  fonts: (string | Record<string, unknown>)[];
  borders: (string | Record<string, unknown>)[];
  fills: (string | Record<string, unknown>)[];
  dxfs: DxfModel[];
}

export interface StylesIndex {
  style?: Record<string, number>;
  numFmt?: Record<string | number, number | string>;
  numFmtNextId?: number;
  font?: Record<string, number>;
  border?: Record<string, number>;
  fill?: Record<string, number>;
  model?: (Record<string, unknown> | undefined)[];
}

// =============================================================================
// StylesXform is used to generate the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
class StylesXform extends BaseXform<StylesModel> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: Record<string, any>;
  index?: StylesIndex;
  weakMap?: WeakMap<object, number>;

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

      // for style manager
      numFmt: new NumFmtXform(),
      font: new FontXform(),
      fill: new FillXform(),
      border: new BorderXform(),
      style: new StyleXform({xfId: true}),

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

    // default (zero) border
    this._addBorder({});

    // add default (all zero) style
    this._addStyle({numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0});

    // add default fills
    this._addFill({type: 'pattern', pattern: 'none'});
    this._addFill({type: 'pattern', pattern: 'gray125'});

    this.weakMap = new WeakMap();
  }

  override render(xmlStream: XmlStreamLike, model?: StylesModel | null): void {
    model = model || this.model!;
    //
    //   <fonts count="2" x14ac:knownFonts="1">
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);

    xmlStream.openNode('styleSheet', StylesXform.STYLESHEET_ATTRIBUTES);

    if (this.index) {
      // model has been built by style manager role (contains xml)
      if (model.numFmts && model.numFmts.length) {
        xmlStream.openNode('numFmts', {count: model.numFmts.length});
        model.numFmts.forEach(numFmtXml => {
          xmlStream.writeXml(numFmtXml as string);
        });
        xmlStream.closeNode();
      }

      if (!model.fonts.length) {
        // default (zero) font
        this._addFont({size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'});
      }
      xmlStream.openNode('fonts', {count: model.fonts.length, 'x14ac:knownFonts': 1});
      model.fonts.forEach(fontXml => {
        xmlStream.writeXml(fontXml as string);
      });
      xmlStream.closeNode();

      xmlStream.openNode('fills', {count: model.fills.length});
      model.fills.forEach(fillXml => {
        xmlStream.writeXml(fillXml as string);
      });
      xmlStream.closeNode();

      xmlStream.openNode('borders', {count: model.borders.length});
      model.borders.forEach(borderXml => {
        xmlStream.writeXml(borderXml as string);
      });
      xmlStream.closeNode();

      this.map.cellStyleXfs.render(xmlStream, [
        {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0},
      ]);

      xmlStream.openNode('cellXfs', {count: model.styles.length});
      model.styles.forEach(styleXml => {
        xmlStream.writeXml(styleXml as string);
      });
      xmlStream.closeNode();
    } else {
      // model is plain JSON and needs to be xformed
      this.map.numFmts.render(xmlStream, model.numFmts);
      this.map.fonts.render(xmlStream, model.fonts);
      this.map.fills.render(xmlStream, model.fills);
      this.map.borders.render(xmlStream, model.borders);
      this.map.cellStyleXfs.render(xmlStream, [
        {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0},
      ]);
      this.map.cellXfs.render(xmlStream, model.styles);
    }

    StylesXform.STATIC_XFORMS.cellStyles.render(xmlStream);

    this.map.dxfs.render(xmlStream, model.dxfs);

    StylesXform.STATIC_XFORMS.tableStyles.render(xmlStream);
    StylesXform.STATIC_XFORMS.extLst.render(xmlStream);

    xmlStream.closeNode();
  }

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(model: Record<string, unknown> | null | undefined, cellType?: number): number {
    if (!model) {
      return 0;
    }

    // if we have no default font, add it here now
    if (!this.model!.fonts.length) {
      // default (zero) font
      this._addFont({size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'});
    }

    // if we have seen this style object before, assume it has the same styleId
    if (this.weakMap && this.weakMap.has(model)) {
      return this.weakMap.get(model)!;
    }

    const style: StyleModel = {};
    cellType = cellType || (Enums as {ValueType: {Number: number; Date: number}}).ValueType.Number;

    if (model.numFmt) {
      style.numFmtId = this._addNumFmtStr(model.numFmt as string);
    } else {
      switch (cellType) {
        case (Enums as {ValueType: {Number: number}}).ValueType.Number:
          style.numFmtId = this._addNumFmtStr('General');
          break;
        case (Enums as {ValueType: {Date: number}}).ValueType.Date:
          style.numFmtId = this._addNumFmtStr('mm-dd-yy');
          break;
        default:
          break;
      }
    }

    if (model.font) {
      style.fontId = this._addFont(model.font as Record<string, unknown>);
    }

    if (model.border) {
      style.borderId = this._addBorder(model.border as Record<string, unknown>);
    }

    if (model.fill) {
      style.fillId = this._addFill(model.fill as Record<string, unknown>);
    }

    if (model.alignment) {
      style.alignment = model.alignment as StyleModel['alignment'];
    }

    if (model.protection) {
      style.protection = model.protection as StyleModel['protection'];
    }

    const styleId = this._addStyle(style);
    if (this.weakMap) {
      this.weakMap.set(model, styleId);
    }
    return styleId;
  }

  addDxfStyle(style: DxfModel): number {
    if (style.numFmt) {
      // register numFmtId to use it during dxf-xform rendering
      style.numFmtId = this._addNumFmtStr(style.numFmt);
    }

    this.model!.dxfs.push(style);
    return this.model!.dxfs.length - 1;
  }

  // =========================================================================
  // Private Interface
  _addStyle(style: StyleModel): number {
    const xml = this.map.style.toXml(style);
    let index = this.index!.style![xml];
    if (index === undefined) {
      index = this.index!.style![xml] = this.model!.styles.length;
      this.model!.styles.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Number Formats
  _addNumFmtStr(formatCode: string): number {
    // check if default format
    let index = NumFmtXform.getDefaultFmtId(formatCode);
    if (index !== undefined) return index;

    // check if already in
    index = this.index!.numFmt![formatCode] as number | undefined;
    if (index !== undefined) return index;

    index = this.index!.numFmt![formatCode] = NUMFMT_BASE + this.model!.numFmts.length;
    const xml = this.map.numFmt.toXml({id: index, formatCode});
    this.model!.numFmts.push(xml);
    return index;
  }

  // =========================================================================
  // Fonts
  _addFont(font: Record<string, unknown>): number {
    const xml = this.map.font.toXml(font);
    let index = this.index!.font![xml];
    if (index === undefined) {
      index = this.index!.font![xml] = this.model!.fonts.length;
      this.model!.fonts.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Borders
  _addBorder(border: Record<string, unknown>): number {
    const xml = this.map.border.toXml(border);
    let index = this.index!.border![xml];
    if (index === undefined) {
      index = this.index!.border![xml] = this.model!.borders.length;
      this.model!.borders.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Fills
  _addFill(fill: Record<string, unknown>): number {
    const xml = this.map.fill.toXml(fill);
    let index = this.index!.fill![xml];
    if (index === undefined) {
      index = this.index!.fill![xml] = this.model!.fills.length;
      this.model!.fills.push(xml);
    }
    return index;
  }

  // =========================================================================

  static STYLESHEET_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x14ac x16r2',
    'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    'xmlns:x16r2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main',
  };

  static STATIC_XFORMS = {
    cellStyles: new StaticXform({
      tag: 'cellStyles',
      $: {count: 1},
      c: [{tag: 'cellStyle', $: {name: 'Normal', xfId: 0, builtinId: 0}}],
    }),
    dxfs: new StaticXform({tag: 'dxfs', $: {count: 0}}),
    tableStyles: new StaticXform({
      tag: 'tableStyles',
      $: {count: 0, defaultTableStyle: 'TableStyleMedium2', defaultPivotStyle: 'PivotStyleLight16'},
    }),
    extLst: new StaticXform({
      tag: 'extLst',
      c: [
        {
          tag: 'ext',
          $: {
            uri: '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}',
            'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
          },
          c: [{tag: 'x14:slicerStyles', $: {defaultSlicerStyle: 'SlicerStyleLight1'}}],
        },
        {
          tag: 'ext',
          $: {
            uri: '{9260A510-F301-46a8-8635-F512D64BE5F5}',
            'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
          },
          c: [{tag: 'x15:timelineStyles', $: {defaultTimelineStyle: 'TimeSlicerStyleLight1'}}],
        },
      ],
    }),
  };
}

export default StylesXform;
export {StylesXform};
