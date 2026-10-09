import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

export interface CfvoModel {
  type: string;
  value?: number;
}

class CfvoXform extends BaseXform<CfvoModel> {
  override tag = 'cfvo';

  override render(xmlStream: XmlStreamLike, model?: CfvoModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.leafNode(this.tag, {
      type: model.type,
      val: model.value,
    });
  }
}

export default CfvoXform;
export {CfvoXform};
