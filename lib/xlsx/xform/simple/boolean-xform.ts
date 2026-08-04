import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface BooleanXformOptions {
  tag: string;
  attr?: string;
}

class BooleanXform extends BaseXform<boolean> {
  attr: string | undefined;

  constructor(options: BooleanXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  override render(xmlStream: XmlStreamLike, model?: boolean | null): void {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  override parseOpen(node: XmlNode): void {
    if (node.name === this.tag) {
      this.model = true;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default BooleanXform;
export {BooleanXform};
