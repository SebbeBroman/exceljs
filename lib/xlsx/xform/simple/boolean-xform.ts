import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default BooleanXform;
export {BooleanXform};
