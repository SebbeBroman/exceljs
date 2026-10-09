import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default PictureXform;
export {PictureXform};
