import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class CNvPicPrXform extends BaseXform {
  override tag = 'xdr:cNvPicPr';

  override render(xmlStream: XmlStreamLike): void {
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('a:picLocks', {
      noChangeAspect: '1',
    });
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        return true;
      default:
        return true;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
}

export default CNvPicPrXform;
export {CNvPicPrXform};
