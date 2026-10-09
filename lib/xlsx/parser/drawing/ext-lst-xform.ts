import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class ExtLstXform extends BaseXform {
  override tag = 'a:extLst';

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
