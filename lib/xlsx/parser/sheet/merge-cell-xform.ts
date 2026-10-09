import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class MergeCellXform extends BaseXform<string> {
  override tag = 'mergeCell';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'mergeCell') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default MergeCellXform;
export {MergeCellXform};
