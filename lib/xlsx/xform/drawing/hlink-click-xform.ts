import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface HLinkClickModel {
  hyperlinks?: {
    rId?: string;
    tooltip?: string;
  };
}

class HLinkClickXform extends BaseXform<HLinkClickModel> {
  override tag = 'a:hlinkClick';

  override render(xmlStream: XmlStreamLike, model?: HLinkClickModel | null): void {
    if (!(model && model.hyperlinks && model.hyperlinks.rId)) {
      return;
    }
    xmlStream.leafNode(this.tag, {
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:id': model.hyperlinks.rId,
      tooltip: model.hyperlinks.tooltip,
    });
  }

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
