import type {StringXformOptions} from '../../xform/simple/string-xform.js';
export type {StringXformOptions} from '../../xform/simple/string-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

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
