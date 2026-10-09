import type {FilterModel} from '../../xform/table/filter-xform.js';
export type {FilterModel} from '../../xform/table/filter-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class FilterXform extends BaseXform<FilterModel> {
  override tag = 'filter';

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
