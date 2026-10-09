import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default HLinkClickXform;
export {HLinkClickXform};
