import type {HyperlinkXformModel} from '../../xform/sheet/hyperlink-xform.js';
export type {HyperlinkXformModel} from '../../xform/sheet/hyperlink-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class HyperlinkXform extends BaseXform<HyperlinkXformModel> {
  override tag = 'hyperlink';

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
