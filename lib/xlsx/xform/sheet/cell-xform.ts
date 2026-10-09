import {dateToExcel} from '../../../utils/utils.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import Range from '../../../model/range.js';
import Enums from '../../../model/enums.js';
import RichTextXform from '../strings/rich-text-xform.js';

export interface CellXformModel {
  address: string;
  type?: number;
  style?: unknown;
  styleId?: number;
  value?: unknown;
  result?: unknown;
  formula?: string;
  sharedFormula?: string;
  shareType?: string;
  si?: number | string;
  ref?: string;
  range?: Range;
  text?: unknown;
  hyperlink?: string;
  tooltip?: string;
  comment?: unknown;
  ssId?: number;
  date1904?: boolean;
  [key: string]: unknown;
}

export interface SharedStringsLike {
  add(value: unknown): number;
  getString(index: number): unknown;
}

export interface StylesCellLike {
  addStyleModel(style: unknown, cellType?: number): number | undefined;
  getStyleModel(styleId: number): {numFmt?: string; [key: string]: unknown} | undefined;
}

export interface MergesLike {
  add(model: CellXformModel): void;
}

export interface CellXformOptions extends XformOptions {
  styles: StylesCellLike;
  sharedStrings?: SharedStringsLike;
  date1904?: boolean;
  comments: unknown[];
  hyperlinks: Array<{address: string; target?: string; tooltip?: string}>;
  merges: MergesLike;
  formulae: Record<string | number, CellXformModel | string>;
  siFormulae: number;
  hyperlinkMap: Record<string, string>;
  commentsMap?: Record<string, unknown>;
}

function getValueType(v: unknown): number {
  if (v === null || v === undefined) {
    return Enums.ValueType.Null;
  }
  if (v instanceof String || typeof v === 'string') {
    return Enums.ValueType.String;
  }
  if (typeof v === 'number') {
    return Enums.ValueType.Number;
  }
  if (typeof v === 'boolean') {
    return Enums.ValueType.Boolean;
  }
  if (v instanceof Date) {
    return Enums.ValueType.Date;
  }
  if (typeof v === 'object' && v !== null) {
    const obj = v as Record<string, unknown>;
    if (obj.text && obj.hyperlink) {
      return Enums.ValueType.Hyperlink;
    }
    if (obj.formula) {
      return Enums.ValueType.Formula;
    }
    if (obj.error) {
      return Enums.ValueType.Error;
    }
  }
  throw new Error('I could not understand type of value');
}

function getEffectiveCellType(cell: CellXformModel): number {
  switch (cell.type) {
    case Enums.ValueType.Formula:
      return getValueType(cell.result);
    default:
      return cell.type as number;
  }
}

class CellXform extends BaseXform<CellXformModel> {
  richTextXForm: RichTextXform;
  t?: string;
  currentNode?: string;

  constructor() {
    super();

    this.richTextXForm = new RichTextXform();
  }

  override tag = 'c';

  override prepare(model?: CellXformModel | null, options?: CellXformOptions): void {
    if (!model || !options) {
      return;
    }
    const styleId = options.styles.addStyleModel(model.style || {}, getEffectiveCellType(model));
    if (styleId) {
      model.styleId = styleId;
    }

    if (model.comment) {
      options.comments.push({...(model.comment as object), ref: model.address});
    }

    switch (model.type) {
      case Enums.ValueType.String:
      case Enums.ValueType.RichText:
        if (options.sharedStrings) {
          model.ssId = options.sharedStrings.add(model.value);
        }
        break;

      case Enums.ValueType.Date:
        if (options.date1904) {
          model.date1904 = true;
        }
        break;

      case Enums.ValueType.Hyperlink:
        if (options.sharedStrings && model.text !== undefined && model.text !== null) {
          model.ssId = options.sharedStrings.add(model.text);
        }
        options.hyperlinks.push({
          address: model.address,
          target: model.hyperlink,
          tooltip: model.tooltip,
        });
        break;

      case Enums.ValueType.Merge:
        options.merges.add(model);
        break;

      case Enums.ValueType.Formula:
        if (options.date1904) {
          // in case valueType is date
          model.date1904 = true;
        }

        if (model.shareType === 'shared') {
          model.si = options.siFormulae++;
        }

        if (model.formula) {
          options.formulae[model.address] = model;
        } else if (model.sharedFormula) {
          const master = options.formulae[model.sharedFormula] as CellXformModel | undefined;
          if (!master || typeof master === 'string') {
            throw new Error(
              `Shared Formula master must exist above and or left of clone for cell ${model.address}`,
            );
          }
          if (master.si === undefined) {
            master.shareType = 'shared';
            master.si = options.siFormulae++;
            master.range = new Range(master.address, model.address);
          } else if (master.range) {
            master.range.expandToAddress(model.address);
          }
          model.si = master.si;
        }
        break;

      default:
        break;
    }
  }

  renderFormula(xmlStream: XmlStreamLike, model: CellXformModel): void {
    let attrs: Record<string, unknown> | null = null;
    switch (model.shareType) {
      case 'shared':
        attrs = {
          t: 'shared',
          ref: model.ref || model.range!.range,
          si: model.si,
        };
        break;

      case 'array':
        attrs = {
          t: 'array',
          ref: model.ref,
        };
        break;

      default:
        if (model.si !== undefined) {
          attrs = {
            t: 'shared',
            si: model.si,
          };
        }
        break;
    }

    switch (getValueType(model.result)) {
      case Enums.ValueType.Null: // ?
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        break;

      case Enums.ValueType.String:
        // oddly, formula results don't ever use shared strings
        xmlStream.addAttribute('t', 'str');
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        xmlStream.leafNode('v', undefined, model.result);
        break;

      case Enums.ValueType.Number:
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        xmlStream.leafNode('v', undefined, model.result);
        break;

      case Enums.ValueType.Boolean:
        xmlStream.addAttribute('t', 'b');
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        xmlStream.leafNode('v', undefined, model.result ? 1 : 0);
        break;

      case Enums.ValueType.Error:
        xmlStream.addAttribute('t', 'e');
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        xmlStream.leafNode('v', undefined, (model.result as {error: string}).error);
        break;

      case Enums.ValueType.Date:
        xmlStream.leafNode('f', attrs as Record<string, unknown> | undefined, model.formula);
        xmlStream.leafNode('v', undefined, dateToExcel(model.result as Date, model.date1904));
        break;

      // case Enums.ValueType.Hyperlink: // ??
      // case Enums.ValueType.Formula:
      default:
        throw new Error('I could not understand type of value');
    }
  }

  override render(xmlStream: XmlStreamLike, model?: CellXformModel | null): void {
    if (!model) {
      return;
    }
    if (model.type === Enums.ValueType.Null && !model.styleId) {
      // if null and no style, exit
      return;
    }

    // Fast path: plain number cell — single XML chunk (address is A1-style, value is numeric)
    if (model.type === Enums.ValueType.Number && !model.styleId) {
      xmlStream.writeXml(`<c r="${model.address}"><v>${model.value as number}</v></c>`);
      return;
    }

    xmlStream.openNode('c');
    xmlStream.addAttribute('r', model.address);

    if (model.styleId) {
      xmlStream.addAttribute('s', model.styleId);
    }

    switch (model.type) {
      case Enums.ValueType.Null:
        break;

      case Enums.ValueType.Number:
        xmlStream.leafNode('v', undefined, model.value);
        break;

      case Enums.ValueType.Boolean:
        xmlStream.addAttribute('t', 'b');
        xmlStream.leafNode('v', undefined, model.value ? '1' : '0');
        break;

      case Enums.ValueType.Error:
        xmlStream.addAttribute('t', 'e');
        xmlStream.leafNode('v', undefined, (model.value as {error: string}).error);
        break;

      case Enums.ValueType.String:
      case Enums.ValueType.RichText:
        if (model.ssId !== undefined) {
          xmlStream.addAttribute('t', 's');
          xmlStream.leafNode('v', undefined, model.ssId);
        } else if (
          model.value &&
          typeof model.value === 'object' &&
          (model.value as {richText?: unknown}).richText
        ) {
          xmlStream.addAttribute('t', 'inlineStr');
          xmlStream.openNode('is');
          (
            model.value as {richText: import('../strings/rich-text-xform.js').RichTextModel[]}
          ).richText.forEach(text => {
            this.richTextXForm.render(xmlStream, text);
          });
          xmlStream.closeNode();
        } else {
          xmlStream.addAttribute('t', 'str');
          xmlStream.leafNode('v', undefined, model.value);
        }
        break;

      case Enums.ValueType.Date:
        xmlStream.leafNode('v', undefined, dateToExcel(model.value as Date, model.date1904));
        break;

      case Enums.ValueType.Hyperlink: {
        // CT_Cell orders children f, v, is — <f> must precede <v> or readers
        // (including our own) mis-parse the cell, and the t attribute belongs on
        // <c> itself, so both are set before the first leafNode.
        //
        // A formula-backed hyperlink must use t="str" with the literal text,
        // never t="s": a shared-string cell carrying <f> reads back as a plain
        // Formula cell and loses the hyperlink.
        const hasFormula = model.formula !== undefined;
        if (hasFormula) {
          xmlStream.addAttribute('t', 'str');
          const fAttrs: Record<string, unknown> = {};
          if (model.shareType === 'shared' && model.si !== undefined) {
            fAttrs.t = 'shared';
            fAttrs.si = model.si;
            if (model.ref !== undefined) fAttrs.ref = model.ref;
          }
          xmlStream.leafNode('f', hasFormula ? fAttrs : undefined, model.formula);
          xmlStream.leafNode('v', undefined, model.text);
        } else if (model.ssId !== undefined) {
          xmlStream.addAttribute('t', 's');
          xmlStream.leafNode('v', undefined, model.ssId);
        } else {
          xmlStream.addAttribute('t', 'str');
          xmlStream.leafNode('v', undefined, model.text);
        }
        break;
      }

      case Enums.ValueType.Formula:
        this.renderFormula(xmlStream, model);
        break;

      case Enums.ValueType.Merge:
        // nothing to add
        break;

      default:
        break;
    }

    xmlStream.closeNode(); // </c>
  }
}

export default CellXform;
export {CellXform};
