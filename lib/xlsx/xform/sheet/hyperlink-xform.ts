import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'hyperlink') {
      this.model = {
        address: node.attributes.ref,
        rId: node.attributes['r:id'],
        tooltip: node.attributes.tooltip,
      };

      // This is an internal link
      if (node.attributes.location) {
        this.model.target = node.attributes.location;
      }
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }

  isInternalLink(model: HyperlinkXformModel): boolean {
    // @example: Sheet2!D3, return true
    return !!(model.target && /^[^!]+![a-zA-Z]+[\d]+$/.test(model.target));
  }
}

export default HyperlinkXform;
export {HyperlinkXform};
