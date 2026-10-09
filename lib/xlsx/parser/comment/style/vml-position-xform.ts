import type {VmlPositionXformOptions} from '../../../xform/comment/style/vml-position-xform.js';
export type {VmlPositionXformOptions} from '../../../xform/comment/style/vml-position-xform.js';
import BaseXform from '../../../base-parser.js';
import type {XmlNode} from '../../../base-parser.js';

class VmlPositionXform extends BaseXform<Record<string, boolean>> {
  _model: VmlPositionXformOptions;
  override tag: string;

  constructor(model: VmlPositionXformOptions) {
    super();
    this._model = model;
    this.tag = model.tag;
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {};
        this.model[this.tag] = true;
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default VmlPositionXform;
export {VmlPositionXform};
