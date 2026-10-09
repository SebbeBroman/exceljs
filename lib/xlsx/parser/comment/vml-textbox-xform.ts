import type {VmlTextboxModel} from '../../xform/comment/vml-textbox-xform.js';
export type {
  VmlTextboxRenderModel,
  VmlTextboxModel,
} from '../../xform/comment/vml-textbox-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class VmlTextboxXform extends BaseXform<VmlTextboxModel> {
  override tag = 'v:textbox';

  reverseConversionUnit(inset: string | undefined): number[] {
    return (inset || '').split(',').map(margin => {
      return Number((parseFloat(margin) * 0.1).toFixed(2));
    });
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          inset: this.reverseConversionUnit(node.attributes.inset),
        };
        return true;
      default:
        return true;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export default VmlTextboxXform;
export {VmlTextboxXform};
