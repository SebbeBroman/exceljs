import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

export interface VmlProtectionXformOptions {
  tag: string;
}

class VmlProtectionXform extends BaseXform {
  _model: VmlProtectionXformOptions;
  override tag: string;
  text?: string;

  constructor(model: VmlProtectionXformOptions) {
    super();
    this._model = model;
    this.tag = model.tag;
  }

  override render(xmlStream: XmlStreamLike, model?: unknown): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }
}

export default VmlProtectionXform;
export {VmlProtectionXform};
