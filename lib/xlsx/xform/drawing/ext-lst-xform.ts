import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class ExtLstXform extends BaseXform {
  override tag = 'a:extLst';

  override render(xmlStream: XmlStreamLike): void {
    xmlStream.openNode(this.tag);
    xmlStream.openNode('a:ext', {
      uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}',
    });
    xmlStream.leafNode('a16:creationId', {
      'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
      id: '{00000000-0008-0000-0000-000002000000}',
    });
    xmlStream.closeNode();
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

export default ExtLstXform;
export {ExtLstXform};
