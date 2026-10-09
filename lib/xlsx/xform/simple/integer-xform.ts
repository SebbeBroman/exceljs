import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface IntegerXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, unknown>;
  zero?: boolean;
}

class IntegerXform extends BaseXform<number> {
  attr: string | undefined;
  attrs: Record<string, unknown> | undefined;
  zero: boolean | undefined;
  text?: string[];

  constructor(options: IntegerXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;

    // option to render zero
    this.zero = options.zero;
  }

  override render(xmlStream: XmlStreamLike, model?: number | null): void {
    // int is different to float in that zero is not rendered
    if (model || this.zero) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, model);
      } else {
        xmlStream.writeText(model);
      }
      xmlStream.closeNode();
    }
  }
}

export default IntegerXform;
export {IntegerXform};
