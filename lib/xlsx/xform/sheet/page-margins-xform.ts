import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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
