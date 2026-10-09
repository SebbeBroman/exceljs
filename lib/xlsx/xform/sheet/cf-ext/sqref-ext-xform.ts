import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

class SqrefExtXform extends BaseXform<string> {
  override tag = 'xm:sqref';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }
}

export default SqrefExtXform;
export {SqrefExtXform};
