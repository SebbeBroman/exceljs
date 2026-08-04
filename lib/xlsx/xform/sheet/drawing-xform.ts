import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface DrawingModel {
  rId: string;
}

class DrawingXform extends BaseXform<DrawingModel> {
  override tag = 'drawing';

  override render(xmlStream: XmlStreamLike, model?: DrawingModel | null): void {
    if (model) {
      xmlStream.leafNode(this.tag, {'r:id': model.rId});
    }
  }

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
