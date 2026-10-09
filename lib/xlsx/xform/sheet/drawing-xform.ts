import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default DrawingXform;
export {DrawingXform};
