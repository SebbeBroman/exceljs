import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface FilterModel {
  val?: string;
}

class FilterXform extends BaseXform<FilterModel> {
  override tag = 'filter';

  override render(xmlStream: XmlStreamLike, model?: FilterModel | null): void {
    xmlStream.leafNode(this.tag, {
      val: model!.val,
    });
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      this.model = {
        val: node.attributes.val,
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

export default FilterXform;
export {FilterXform};
