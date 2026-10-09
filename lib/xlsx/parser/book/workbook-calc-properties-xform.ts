import type {WorkbookCalcPropertiesModel} from '../../xform/book/workbook-calc-properties-xform.js';
export type {WorkbookCalcPropertiesModel} from '../../xform/book/workbook-calc-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class WorkbookCalcPropertiesXform extends BaseXform<WorkbookCalcPropertiesModel> {
  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'calcPr') {
      this.model = {};
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default WorkbookCalcPropertiesXform;
export {WorkbookCalcPropertiesXform};
