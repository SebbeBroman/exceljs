import type {PageSetupPropertiesModel} from '../../xform/sheet/page-setup-properties-xform.js';
export type {PageSetupPropertiesModel} from '../../xform/sheet/page-setup-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class PageSetupPropertiesXform extends BaseXform<PageSetupPropertiesModel> {
  override tag = 'pageSetUpPr';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      this.model = {
        fitToPage: node.attributes.fitToPage === '1',
      };
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default PageSetupPropertiesXform;
export {PageSetupPropertiesXform};
