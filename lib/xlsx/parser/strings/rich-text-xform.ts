import type {RichTextModel} from '../../xform/strings/rich-text-xform.js';
export type {RichTextModel} from '../../xform/strings/rich-text-xform.js';
import TextXform from './text-xform.js';
import FontXform from '../style/font-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class RichTextXform extends BaseXform<RichTextModel> {
  _textXform?: TextXform;
  _fontXform?: FontXform;

  constructor(model?: RichTextModel) {
    super();

    this.model = model;
  }

  override tag = 'r';

  get textXform(): TextXform {
    return this._textXform || (this._textXform = new TextXform());
  }

  get fontXform(): FontXform {
    return this._fontXform || (this._fontXform = new FontXform(RichTextXform.FONT_OPTIONS));
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'r':
        this.model = {};
        return true;
      case 't':
        this.parser = this.textXform;
        this.parser.parseOpen(node);
        return true;
      case 'rPr':
        this.parser = this.fontXform;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    switch (name) {
      case 'r':
        return false;
      case 't':
        (this.model as RichTextModel).text = this.parser!.model as string;
        this.parser = undefined;
        return true;
      case 'rPr':
        (this.model as RichTextModel).font = this.parser!.model;
        this.parser = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
        }
        return true;
    }
  }

  static FONT_OPTIONS = {
    tagName: 'rPr',
    fontNameTag: 'rFont',
  };
}

export default RichTextXform;
export {RichTextXform};
