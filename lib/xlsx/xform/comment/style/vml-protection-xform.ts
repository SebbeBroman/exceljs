import BaseXform from '../../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';

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
