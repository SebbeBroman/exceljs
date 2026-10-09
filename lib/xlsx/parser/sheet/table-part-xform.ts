import type {TablePartModel} from '../../xform/sheet/table-part-xform.js';
export type {TablePartModel} from '../../xform/sheet/table-part-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class TablePartXform extends BaseXform<TablePartModel> {
  override tag = 'tablePart';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:id'],
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

export default TablePartXform;
export {TablePartXform};
