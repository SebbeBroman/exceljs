import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import BlipXform from './blip-xform.js';
import type {BlipModel} from './blip-xform.js';

class BlipFillXform extends BaseXform<BlipModel> {
  override tag = 'xdr:blipFill';
  declare map: {
    'a:blip': BlipXform;
  };

  constructor() {
    super();

    this.map = {
      'a:blip': new BlipXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: BlipModel | null): void {
    xmlStream.openNode(this.tag);

    this.map['a:blip'].render(xmlStream, model);

    // TODO: options for this + parsing
    xmlStream.openNode('a:stretch');
    xmlStream.leafNode('a:fillRect');
    xmlStream.closeNode();

    xmlStream.closeNode();
  }
}

export default BlipFillXform;
export {BlipFillXform};
