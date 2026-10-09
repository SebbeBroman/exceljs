import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface PageMarginsModel {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

class PageMarginsXform extends BaseXform<PageMarginsModel> {
  override tag = 'pageMargins';

  override render(xmlStream: XmlStreamLike, model?: Partial<PageMarginsModel> | null): void {
    if (model) {
      const attributes = {
        left: model.left,
        right: model.right,
        top: model.top,
        bottom: model.bottom,
        header: model.header,
        footer: model.footer,
      };
      if (Object.values(attributes).some(value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }
}

export default PageMarginsXform;
export {PageMarginsXform};
