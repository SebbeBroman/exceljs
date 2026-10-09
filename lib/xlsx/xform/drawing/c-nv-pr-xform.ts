import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import HlickClickXform from './hlink-click-xform.js';
import type {HLinkClickModel} from './hlink-click-xform.js';
import ExtLstXform from './ext-lst-xform.js';

export interface CNvPrModel extends HLinkClickModel {
  index?: number;
}

class CNvPrXform extends BaseXform<HLinkClickModel> {
  override tag = 'xdr:cNvPr';
  declare map: {
    'a:hlinkClick': HlickClickXform;
    'a:extLst': ExtLstXform;
  };

  constructor() {
    super();

    this.map = {
      'a:hlinkClick': new HlickClickXform(),
      'a:extLst': new ExtLstXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CNvPrModel | null): void {
    xmlStream.openNode(this.tag, {
      id: model!.index,
      name: `Picture ${model!.index}`,
    });
    this.map['a:hlinkClick'].render(xmlStream, model);
    this.map['a:extLst'].render(xmlStream);
    xmlStream.closeNode();
  }
}

export default CNvPrXform;
export {CNvPrXform};
