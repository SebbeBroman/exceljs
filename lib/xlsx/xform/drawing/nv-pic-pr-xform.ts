import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import CNvPrXform from './c-nv-pr-xform.js';
import type {CNvPrModel} from './c-nv-pr-xform.js';
import CNvPicPrXform from './c-nv-pic-pr-xform.js';
import type {HLinkClickModel} from './hlink-click-xform.js';

class NvPicPrXform extends BaseXform<HLinkClickModel> {
  override tag = 'xdr:nvPicPr';
  declare map: {
    'xdr:cNvPr': CNvPrXform;
    'xdr:cNvPicPr': CNvPicPrXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:cNvPr': new CNvPrXform(),
      'xdr:cNvPicPr': new CNvPicPrXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CNvPrModel | null): void {
    xmlStream.openNode(this.tag);
    this.map['xdr:cNvPr'].render(xmlStream, model);
    this.map['xdr:cNvPicPr'].render(xmlStream);
    xmlStream.closeNode();
  }
}

export default NvPicPrXform;
export {NvPicPrXform};
