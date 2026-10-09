export type {
  VmlAnchorRenderModel,
  VmlRefAddress,
  VmlAnchorBox,
} from '../../xform/comment/vml-anchor-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

// render the triangle in the cell for the comment
class VmlAnchorXform extends BaseXform {
  override tag = 'x:Anchor';
  text?: string;

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

export default VmlAnchorXform;
export {VmlAnchorXform};
