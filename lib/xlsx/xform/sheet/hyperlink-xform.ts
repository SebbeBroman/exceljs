import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface HyperlinkXformModel {
  address: string;
  rId?: string;
  tooltip?: string;
  target?: string;
}

class HyperlinkXform extends BaseXform<HyperlinkXformModel> {
  override tag = 'hyperlink';

  override render(xmlStream: XmlStreamLike, model?: HyperlinkXformModel | null): void {
    if (!model) {
      return;
    }
    if (this.isInternalLink(model)) {
      xmlStream.leafNode('hyperlink', {
        ref: model.address,
        'r:id': model.rId,
        tooltip: model.tooltip,
        location: model.target,
      });
    } else {
      xmlStream.leafNode('hyperlink', {
        ref: model.address,
        'r:id': model.rId,
        tooltip: model.tooltip,
      });
    }
  }

  isInternalLink(model: HyperlinkXformModel): boolean {
    // @example: Sheet2!D3, return true
    return !!(model.target && /^[^!]+![a-zA-Z]+[\d]+$/.test(model.target));
  }
}

export default HyperlinkXform;
export {HyperlinkXform};
