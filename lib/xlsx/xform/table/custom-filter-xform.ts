import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface CustomFilterModel {
  val?: string;
  operator?: string;
}

class CustomFilterXform extends BaseXform<CustomFilterModel> {
  override tag = 'customFilter';

  override render(xmlStream: XmlStreamLike, model?: CustomFilterModel | null): void {
    xmlStream.leafNode(this.tag, {
      val: model!.val,
      operator: model!.operator,
    });
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      this.model = {
        val: node.attributes.val,
        operator: node.attributes.operator,
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

export default CustomFilterXform;
export {CustomFilterXform};
