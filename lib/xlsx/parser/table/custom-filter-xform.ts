import type {CustomFilterModel} from '../../xform/table/custom-filter-xform.js';
export type {CustomFilterModel} from '../../xform/table/custom-filter-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class CustomFilterXform extends BaseXform<CustomFilterModel> {
  override tag = 'customFilter';

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
