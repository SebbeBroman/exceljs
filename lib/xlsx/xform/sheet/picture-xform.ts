import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface PictureModel {
  rId: string;
}

class PictureXform extends BaseXform<PictureModel> {
  override tag = 'picture';

  override render(xmlStream: XmlStreamLike, model?: PictureModel | null): void {
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

export default PictureXform;
export {PictureXform};
