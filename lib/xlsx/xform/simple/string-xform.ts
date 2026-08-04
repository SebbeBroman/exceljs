import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface StringXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, unknown>;
}

class StringXform extends BaseXform<string> {
  attr: string | undefined;
  attrs: Record<string, unknown> | undefined;
  text?: string[];

  constructor(options: StringXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;
  }

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    if (model !== undefined && model !== null) {
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

  override parseOpen(node: XmlNode): void {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = node.attributes[this.attr];
      } else {
        this.text = [];
      }
    }
  }

  override parseText(text: string): void {
    if (!this.attr) {
      this.text!.push(text);
    }
  }

  override parseClose(): boolean {
    if (!this.attr) {
      this.model = this.text!.join('');
    }
    return false;
  }
}

export default StringXform;
export {StringXform};
