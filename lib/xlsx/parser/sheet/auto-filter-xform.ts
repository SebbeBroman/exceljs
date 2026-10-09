export type {
  AutoFilterModel,
  AutoFilterRangeModel,
  AutoFilterAddress,
} from '../../xform/sheet/auto-filter-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class AutoFilterXform extends BaseXform<string> {
  override tag = 'autoFilter';

  override parseOpen(node: XmlNode): void {
    if (node.name === 'autoFilter') {
      this.model = node.attributes.ref;
    }
  }
}

export default AutoFilterXform;
export {AutoFilterXform};
