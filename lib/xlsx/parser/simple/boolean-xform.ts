import type {BooleanXformOptions} from '../../xform/simple/boolean-xform.js';
export type {BooleanXformOptions} from '../../xform/simple/boolean-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class BooleanXform extends BaseXform<boolean> {
  attr: string | undefined;

  constructor(options: BooleanXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
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
