import TextXform from './text-xform.js';
import FontXform from '../style/font-xform.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

// <r>
//   <rPr>
//     <sz val="11"/>
//     <color theme="1" tint="5"/>
//     <rFont val="Calibri"/>
//     <family val="2"/>
//     <scheme val="minor"/>
//   </rPr>
//   <t xml:space="preserve"> is </t>
// </r>

export interface RichTextModel {
  font?: unknown;
  text?: string;
}

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

  override render(xmlStream: XmlStreamLike, model?: RichTextModel | null): void {
    model = model || this.model;

    xmlStream.openNode('r');
    if (model?.font) {
      this.fontXform.render(xmlStream, model.font as never);
    }
    this.textXform.render(xmlStream, model?.text);
    xmlStream.closeNode();
  }

  static FONT_OPTIONS = {
    tagName: 'rPr',
    fontNameTag: 'rFont',
  };
}

export default RichTextXform;
export {RichTextXform};
