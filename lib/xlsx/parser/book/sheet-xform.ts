import type {SheetModel} from '../../xform/book/sheet-xform.js';
export type {SheetModel} from '../../xform/book/sheet-xform.js';
import {xmlDecode} from '../../../utils/utils.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class WorksheetXform extends BaseXform<SheetModel> {
  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'sheet') {
      this.model = {
        name: xmlDecode(node.attributes.name),
        id: parseInt(node.attributes.sheetId, 10),
        state: node.attributes.state,
        rId: node.attributes['r:id'],
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

export default WorksheetXform;
export {WorksheetXform};
