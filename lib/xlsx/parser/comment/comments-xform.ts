import type {CommentsModel} from '../../xform/comment/comments-xform.js';
export type {CommentsModel} from '../../xform/comment/comments-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import CommentXform from './comment-xform.js';
import type {CommentXformModel} from './comment-xform.js';

class CommentsXform extends BaseXform<CommentsModel> {
  declare map: {
    comment: CommentXform;
  };

  constructor() {
    super();
    this.map = {
      comment: new CommentXform(),
    };
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'commentList':
        this.model = {
          comments: [],
        };
        return true;
      case 'comment':
        this.parser = this.map.comment;
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
      case 'commentList':
        return false;
      case 'comment':
        this.model!.comments.push(this.parser!.model as CommentXformModel);
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

export default CommentsXform;
export {CommentsXform};
