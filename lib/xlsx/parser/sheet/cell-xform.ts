import type {CellXformOptions, CellXformModel} from '../../xform/sheet/cell-xform.js';
export type {
  CellXformOptions,
  MergesLike,
  StylesCellLike,
  SharedStringsLike,
  CellXformModel,
} from '../../xform/sheet/cell-xform.js';
import {excelToDate, isDateFmt, xmlDecode} from '../../../utils/utils.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import Enums from '../../../model/enums.js';
import RichTextXform from '../strings/rich-text-xform.js';

class CellXform extends BaseXform<CellXformModel> {
  richTextXForm: RichTextXform;
  t?: string;
  currentNode?: string;

  constructor() {
    super();

    this.richTextXForm = new RichTextXform();
  }

  override tag = 'c';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'c':
        // const address = colCache.decodeAddress(node.attributes.r);
        this.model = {
          address: node.attributes.r,
        };
        this.t = node.attributes.t;
        if (node.attributes.s) {
          this.model.styleId = parseInt(node.attributes.s, 10);
        }
        return true;

      case 'f':
        this.currentNode = 'f';
        this.model!.si = node.attributes.si;
        this.model!.shareType = node.attributes.t;
        this.model!.ref = node.attributes.ref;
        return true;

      case 'v':
        this.currentNode = 'v';
        return true;

      case 't':
        this.currentNode = 't';
        return true;

      case 'r':
        this.parser = this.richTextXForm;
        this.parser.parseOpen(node);
        return true;

      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
      return;
    }
    switch (this.currentNode) {
      case 'f':
        this.model!.formula = this.model!.formula ? this.model!.formula + text : text;
        break;
      case 'v':
      case 't':
        if (
          this.model!.value &&
          typeof this.model!.value === 'object' &&
          (this.model!.value as {richText?: {text?: string}}).richText
        ) {
          const richText = (this.model!.value as {richText: {text?: string}}).richText;
          richText.text = richText.text ? richText.text + text : text;
        } else {
          this.model!.value = this.model!.value ? (this.model!.value as string) + text : text;
        }
        break;
      default:
        break;
    }
  }

  override parseClose(name?: string): boolean {
    switch (name) {
      case 'c': {
        const {model} = this;
        if (!model) {
          return false;
        }

        // first guess on cell type
        if (model.formula || model.shareType) {
          model.type = Enums.ValueType.Formula;
          if (model.value) {
            if (this.t === 'str') {
              model.result = xmlDecode(model.value as string);
            } else if (this.t === 'b') {
              model.result = parseInt(model.value as string, 10) !== 0;
            } else if (this.t === 'e') {
              model.result = {error: model.value};
            } else {
              model.result = parseFloat(model.value as string);
            }
            model.value = undefined;
          }
        } else if (model.value !== undefined) {
          switch (this.t) {
            case 's':
              model.type = Enums.ValueType.String;
              model.value = parseInt(model.value as string, 10);
              break;
            case 'str':
              model.type = Enums.ValueType.String;
              model.value = xmlDecode(model.value as string);
              break;
            case 'inlineStr':
              model.type = Enums.ValueType.String;
              break;
            case 'b':
              model.type = Enums.ValueType.Boolean;
              model.value = parseInt(model.value as string, 10) !== 0;
              break;
            case 'e':
              model.type = Enums.ValueType.Error;
              model.value = {error: model.value};
              break;
            default:
              model.type = Enums.ValueType.Number;
              model.value = parseFloat(model.value as string);
              break;
          }
        } else if (model.styleId) {
          model.type = Enums.ValueType.Null;
        } else {
          model.type = Enums.ValueType.Merge;
        }
        return false;
      }

      case 'f':
      case 'v':
      case 'is':
        this.currentNode = undefined;
        return true;

      case 't':
        if (this.parser) {
          this.parser.parseClose(name);
          return true;
        }
        this.currentNode = undefined;
        return true;

      case 'r':
        this.model!.value = this.model!.value || {};
        {
          const value = this.model!.value as {richText?: unknown[]};
          value.richText = value.richText || [];
          value.richText.push(this.parser!.model);
        }
        this.parser = undefined;
        this.currentNode = undefined;
        return true;

      default:
        if (this.parser) {
          this.parser.parseClose(name);
          return true;
        }
        return false;
    }
  }

  override reconcile(model?: CellXformModel | null, options?: CellXformOptions): void {
    if (!model || !options) {
      return;
    }
    const style =
      model.styleId !== undefined && options.styles
        ? options.styles.getStyleModel(model.styleId)
        : undefined;
    if (style) {
      model.style = style;
    }
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    switch (model.type) {
      case Enums.ValueType.String:
        if (typeof model.value === 'number') {
          if (options.sharedStrings) {
            model.value = options.sharedStrings.getString(model.value);
          }
        }
        if (
          model.value &&
          typeof model.value === 'object' &&
          (model.value as {richText?: unknown}).richText
        ) {
          model.type = Enums.ValueType.RichText;
        }
        break;

      case Enums.ValueType.Number:
        if (style && isDateFmt(style.numFmt)) {
          model.type = Enums.ValueType.Date;
          model.value = excelToDate(model.value as number, options.date1904);
        }
        break;

      case Enums.ValueType.Formula:
        if (model.result !== undefined && style && isDateFmt(style.numFmt)) {
          model.result = excelToDate(model.result as number, options.date1904);
        }
        if (model.shareType === 'shared') {
          if (model.ref) {
            // master
            options.formulae[model.si as number] = model.address;
          } else {
            // slave
            model.sharedFormula = options.formulae[model.si as number] as string;
            delete model.shareType;
          }
          delete model.si;
        }
        break;

      default:
        break;
    }

    // look for hyperlink
    const hyperlink = options.hyperlinkMap[model.address];
    if (hyperlink) {
      if (model.type === Enums.ValueType.Formula) {
        model.text = model.result;
        model.result = undefined;
      } else {
        model.text = model.value;
        model.value = undefined;
      }
      model.type = Enums.ValueType.Hyperlink;
      model.hyperlink = hyperlink;
    }

    const comment = options.commentsMap && options.commentsMap[model.address];
    if (comment) {
      model.comment = comment;
    }
  }
}

export default CellXform;
export {CellXform};
