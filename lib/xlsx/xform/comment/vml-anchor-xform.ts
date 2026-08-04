import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface VmlAnchorBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface VmlRefAddress {
  col: number;
  row: number;
}

export interface VmlAnchorRenderModel {
  anchor?: VmlAnchorBox;
  refAddress?: VmlRefAddress;
}

// render the triangle in the cell for the comment
class VmlAnchorXform extends BaseXform {
  override tag = 'x:Anchor';
  text?: string;

  getAnchorRect(anchor: VmlAnchorBox): number[] {
    const l = Math.floor(anchor.left);
    const lf = Math.floor((anchor.left - l) * 68);
    const t = Math.floor(anchor.top);
    const tf = Math.floor((anchor.top - t) * 18);
    const r = Math.floor(anchor.right);
    const rf = Math.floor((anchor.right - r) * 68);
    const b = Math.floor(anchor.bottom);
    const bf = Math.floor((anchor.bottom - b) * 18);
    return [l, lf, t, tf, r, rf, b, bf];
  }

  getDefaultRect(ref: VmlRefAddress): number[] {
    const l = ref.col;
    const lf = 6;
    const t = Math.max(ref.row - 2, 0);
    const tf = 14;
    const r = l + 2;
    const rf = 2;
    const b = t + 4;
    const bf = 16;
    return [l, lf, t, tf, r, rf, b, bf];
  }

  override render(xmlStream: XmlStreamLike, model?: VmlAnchorRenderModel | null): void {
    const rect = model!.anchor
      ? this.getAnchorRect(model!.anchor)
      : this.getDefaultRect(model!.refAddress!);

    xmlStream.leafNode('x:Anchor', undefined, rect.join(', '));
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

export default VmlAnchorXform;
export {VmlAnchorXform};
