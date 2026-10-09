import type {IntegerXformOptions} from '../../xform/simple/integer-xform.js';
export type {IntegerXformOptions} from '../../xform/simple/integer-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

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
