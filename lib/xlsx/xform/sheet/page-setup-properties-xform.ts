import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface PageSetupPropertiesModel {
  fitToPage?: boolean;
}

class PageSetupPropertiesXform extends BaseXform<PageSetupPropertiesModel> {
  override tag = 'pageSetUpPr';

  override render(xmlStream: XmlStreamLike, model?: PageSetupPropertiesModel | null): boolean {
    if (model && model.fitToPage) {
      xmlStream.leafNode(this.tag, {
        fitToPage: model.fitToPage ? '1' : undefined,
      });
      return true;
    }
    return false;
  }

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
