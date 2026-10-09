import type {DrawingModel} from '../../xform/sheet/drawing-xform.js';
export type {DrawingModel} from '../../xform/sheet/drawing-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class DrawingXform extends BaseXform<DrawingModel> {
  override tag = 'drawing';

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

export default DrawingXform;
export {DrawingXform};
