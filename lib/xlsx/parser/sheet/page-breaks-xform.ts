import type {PageBreakModel} from '../../xform/sheet/page-breaks-xform.js';
export type {PageBreakModel} from '../../xform/sheet/page-breaks-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class PageBreaksXform extends BaseXform<PageBreakModel | string> {
  override tag = 'brk';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'brk') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default PageBreaksXform;
export {PageBreaksXform};
