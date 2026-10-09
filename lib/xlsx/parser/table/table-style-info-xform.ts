import type {TableStyleInfoModel} from '../../xform/table/table-style-info-xform.js';
export type {TableStyleInfoModel} from '../../xform/table/table-style-info-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class TableStyleInfoXform extends BaseXform<TableStyleInfoModel> {
  override tag = 'tableStyleInfo';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        theme: attributes.name ? attributes.name : null,
        showFirstColumn: attributes.showFirstColumn === '1',
        showLastColumn: attributes.showLastColumn === '1',
        showRowStripes: attributes.showRowStripes === '1',
        showColumnStripes: attributes.showColumnStripes === '1',
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

export default TableStyleInfoXform;
export {TableStyleInfoXform};
