import type {TableColumnModel} from '../../xform/table/table-column-xform.js';
export type {TableColumnModel} from '../../xform/table/table-column-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class TableColumnXform extends BaseXform<TableColumnModel> {
  override tag = 'tableColumn';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        name: attributes.name,
        totalsRowLabel: attributes.totalsRowLabel,
        totalsRowFunction: attributes.totalsRowFunction,
        dxfId: attributes.dxfId,
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

export default TableColumnXform;
export {TableColumnXform};
