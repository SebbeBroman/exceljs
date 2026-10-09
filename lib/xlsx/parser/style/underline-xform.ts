import type {UnderlineModel} from '../../xform/style/underline-xform.js';
export type {UnderlineModel} from '../../xform/style/underline-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class UnderlineXform extends BaseXform<UnderlineModel> {
  constructor(model?: UnderlineModel) {
    super();

    this.model = model;
  }

  override tag = 'u';

  override parseOpen(node: XmlNode): void {
    if (node.name === 'u') {
      this.model = node.attributes.val || true;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default UnderlineXform;
export {UnderlineXform};
