import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

export interface CfIconExtModel {
  iconSet?: string;
  iconId?: number;
}

class CfIconExtXform extends BaseXform<CfIconExtModel> {
  override tag = 'x14:cfIcon';

  override render(xmlStream: XmlStreamLike, model?: CfIconExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.leafNode(this.tag, {
      iconSet: model.iconSet,
      iconId: model.iconId,
    });
  }
}

export default CfIconExtXform;
export {CfIconExtXform};
