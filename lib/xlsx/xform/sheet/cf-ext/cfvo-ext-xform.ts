import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';
import FExtXform from './f-ext-xform.js';

export interface CfvoExtModel {
  type?: string;
  value?: number;
}

class CfvoExtXform extends CompositeXform<CfvoExtModel> {
  fExtXform: FExtXform;

  constructor() {
    super();

    this.map = {
      'xm:f': (this.fExtXform = new FExtXform()),
    };
  }

  override tag = 'x14:cfvo';

  override render(xmlStream: XmlStreamLike, model?: CfvoExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag, {
      type: model.type,
    });
    if (model.value !== undefined) {
      this.fExtXform.render(xmlStream, model.value);
    }
    xmlStream.closeNode();
  }
}

export default CfvoExtXform;
export {CfvoExtXform};
