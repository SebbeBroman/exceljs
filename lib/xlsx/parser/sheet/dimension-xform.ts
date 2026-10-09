import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class DimensionXform extends BaseXform<string> {
  override tag = 'dimension';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'dimension') {
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

export default DimensionXform;
export {DimensionXform};
