import type {HLinkClickModel} from '../../xform/drawing/hlink-click-xform.js';
export type {HLinkClickModel} from '../../xform/drawing/hlink-click-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class HLinkClickXform extends BaseXform<HLinkClickModel> {
  override tag = 'a:hlinkClick';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          hyperlinks: {
            rId: node.attributes['r:id'],
            tooltip: node.attributes.tooltip,
          },
        };
        return true;
      default:
        return true;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default HLinkClickXform;
export {HLinkClickXform};
