import type {CommentXformModel, CommentNoteText} from '../../xform/comment/comment-xform.js';
export type {CommentXformModel, CommentNoteText} from '../../xform/comment/comment-xform.js';
import RichTextXform from '../strings/rich-text-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'comment':
        this.model = {
          type: 'note',
          note: {
            texts: [],
          },
          ...node.attributes,
        };
        return true;
      case 'r':
        this.parser = this.richTextXform;
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
      case 'comment':
        return false;
      case 'r':
        this.model!.note!.texts!.push(this.parser!.model as CommentNoteText);
        this.parser = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
        }
        return true;
    }
  }
}

export default CommentXform;
export {CommentXform};
