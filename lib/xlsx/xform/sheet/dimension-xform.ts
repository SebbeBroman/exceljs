import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class DimensionXform extends BaseXform<string> {
  override tag = 'dimension';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    if (model) {
      xmlStream.leafNode('dimension', {ref: model});
    }
  }

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
