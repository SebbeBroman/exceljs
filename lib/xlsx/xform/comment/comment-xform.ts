import RichTextXform from '../strings/rich-text-xform.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

/**
  <comment ref="B1" authorId="0">
    <text>
      <r>
        <rPr>
          <b/>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t>51422:</t>
      </r>
      <r>
        <rPr>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t xml:space="preserve">&#10;test</t>
      </r>
    </text>
  </comment>
 */

export interface CommentNoteText {
  text?: string;
  font?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface CommentXformModel {
  ref?: string;
  type?: string;
  note?: {
    texts?: CommentNoteText[];
  };
  authorId?: string | number;
  [key: string]: unknown;
}

class CommentXform extends BaseXform<CommentXformModel> {
  override tag = 'r';
  _richTextXform?: RichTextXform;

  constructor(model?: CommentXformModel) {
    super();
    this.model = model;
  }

  get richTextXform(): RichTextXform {
    if (!this._richTextXform) {
      this._richTextXform = new RichTextXform();
    }
    return this._richTextXform;
  }

  override render(xmlStream: XmlStreamLike, model?: CommentXformModel | null): void {
    model = model || this.model;

    xmlStream.openNode('comment', {
      ref: model!.ref,
      authorId: 0,
    });
    xmlStream.openNode('text');
    if (model && model.note && model.note.texts) {
      model.note.texts.forEach(text => {
        this.richTextXform.render(xmlStream, text);
      });
    }
    xmlStream.closeNode();
    xmlStream.closeNode();
  }
}

export default CommentXform;
export {CommentXform};
