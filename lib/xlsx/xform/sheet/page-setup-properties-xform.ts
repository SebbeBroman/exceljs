import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default PageSetupPropertiesXform;
export {PageSetupPropertiesXform};
