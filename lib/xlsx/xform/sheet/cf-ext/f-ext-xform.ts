import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

class FExtXform extends BaseXform<string> {
  override tag = 'xm:f';

  override render(xmlStream: XmlStreamLike, model?: string | number | null): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }
}

export default FExtXform;
export {FExtXform};
