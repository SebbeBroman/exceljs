import type {PictureModel} from '../../xform/sheet/picture-xform.js';
export type {PictureModel} from '../../xform/sheet/picture-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class PictureXform extends BaseXform<PictureModel> {
  override tag = 'picture';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:id'],
        };
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default PictureXform;
export {PictureXform};
