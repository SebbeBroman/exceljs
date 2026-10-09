import type {VmlProtectionXformOptions} from '../../../xform/comment/style/vml-protection-xform.js';
export type {VmlProtectionXformOptions} from '../../../xform/comment/style/vml-protection-xform.js';
import BaseXform from '../../../base-parser.js';
import type {XmlNode} from '../../../base-parser.js';

class VmlProtectionXform extends BaseXform {
  _model: VmlProtectionXformOptions;
  override tag: string;
  text?: string;

  constructor(model: VmlProtectionXformOptions) {
    super();
    this._model = model;
    this.tag = model.tag;
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.text = '';
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    this.text = text;
  }

  override parseClose(): boolean {
    return false;
  }
}

export default VmlProtectionXform;
export {VmlProtectionXform};
