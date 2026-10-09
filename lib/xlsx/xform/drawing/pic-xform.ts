import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import StaticXform from '../static-xform.js';
import BlipFillXform from './blip-fill-xform.js';
import NvPicPrXform from './nv-pic-pr-xform.js';
import spPrJSON from './sp-pr.js';
import type {HLinkClickModel} from './hlink-click-xform.js';

export interface PicModel extends HLinkClickModel {
  rId?: string;
  index?: number;
}

class PicXform extends BaseXform<PicModel> {
  override tag = 'xdr:pic';
  declare map: {
    'xdr:nvPicPr': NvPicPrXform;
    'xdr:blipFill': BlipFillXform;
    'xdr:spPr': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:nvPicPr': new NvPicPrXform(),
      'xdr:blipFill': new BlipFillXform(),
      'xdr:spPr': new StaticXform(spPrJSON),
    };
  }

  override prepare(model?: PicModel | null, options?: XformOptions): void {
    model!.index = (options!.index as number) + 1;
  }

  override render(xmlStream: XmlStreamLike, model?: PicModel | null): void {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvPicPr'].render(xmlStream, model);
    this.map['xdr:blipFill'].render(xmlStream, model as {rId: string});
    this.map['xdr:spPr'].render(xmlStream);

    xmlStream.closeNode();
  }
}

export default PicXform;
export {PicXform};
