import type {BlipModel} from '../../xform/drawing/blip-xform.js';
export type {BlipModel} from '../../xform/drawing/blip-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class BlipXform extends BaseXform<BlipModel> {
  override tag = 'a:blip';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:embed'],
        };
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

export default BlipXform;
export {BlipXform};
