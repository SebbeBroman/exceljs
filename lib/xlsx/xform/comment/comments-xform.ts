import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import CommentXform from './comment-xform.js';
import type {CommentXformModel} from './comment-xform.js';

export interface CommentsModel {
  comments: CommentXformModel[];
}

class CommentsXform extends BaseXform<CommentsModel> {
  static COMMENTS_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  };

  declare map: {
    comment: CommentXform;
  };

  constructor() {
    super();
    this.map = {
      comment: new CommentXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CommentsModel | null): void {
    model = model || this.model;
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('comments', CommentsXform.COMMENTS_ATTRIBUTES);

    // authors
    // TODO: support authors properly
    xmlStream.openNode('authors');
    xmlStream.leafNode('author', undefined, 'Author');
    xmlStream.closeNode();

    // comments
    xmlStream.openNode('commentList');
    model!.comments.forEach(comment => {
      this.map.comment.render(xmlStream, comment);
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }
}

export default CommentsXform;
export {CommentsXform};
