import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

class CNvPicPrXform extends BaseXform {
  override tag = 'xdr:cNvPicPr';

  override render(xmlStream: XmlStreamLike): void {
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('a:picLocks', {
      noChangeAspect: '1',
    });
    xmlStream.closeNode();
  }
}

export default CNvPicPrXform;
export {CNvPicPrXform};
