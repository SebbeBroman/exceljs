import type {WorkbookPropertiesModel} from '../../xform/book/workbook-properties-xform.js';
export type {WorkbookPropertiesModel} from '../../xform/book/workbook-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class WorksheetPropertiesXform extends BaseXform<WorkbookPropertiesModel> {
  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'workbookPr') {
      this.model = {
        date1904: node.attributes.date1904 === '1',
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

export default WorksheetPropertiesXform;
export {WorksheetPropertiesXform};
