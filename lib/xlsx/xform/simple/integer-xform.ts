import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = parseInt(node.attributes[this.attr], 10);
      } else {
        this.text = [];
      }
      return true;
    }
    return false;
  }

  override parseText(text: string): void {
    if (!this.attr) {
      this.text!.push(text);
    }
  }

  override parseClose(): boolean {
    if (!this.attr) {
      this.model = parseInt(this.text!.join('') || '0', 10);
    }
    return false;
  }
}

export default IntegerXform;
export {IntegerXform};
