import type {PageMarginsModel} from '../../xform/sheet/page-margins-xform.js';
export type {PageMarginsModel} from '../../xform/sheet/page-margins-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class PageMarginsXform extends BaseXform<PageMarginsModel> {
  override tag = 'pageMargins';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          left: parseFloat(node.attributes.left || (0.7 as unknown as string)),
          right: parseFloat(node.attributes.right || (0.7 as unknown as string)),
          top: parseFloat(node.attributes.top || (0.75 as unknown as string)),
          bottom: parseFloat(node.attributes.bottom || (0.75 as unknown as string)),
          header: parseFloat(node.attributes.header || (0.3 as unknown as string)),
          footer: parseFloat(node.attributes.footer || (0.3 as unknown as string)),
        };
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default PageMarginsXform;
export {PageMarginsXform};
