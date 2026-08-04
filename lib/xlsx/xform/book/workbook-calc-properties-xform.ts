import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface WorkbookCalcPropertiesModel {
  fullCalcOnLoad?: boolean;
}

class WorkbookCalcPropertiesXform extends BaseXform<WorkbookCalcPropertiesModel> {
  override render(xmlStream: XmlStreamLike, model?: WorkbookCalcPropertiesModel | null): void {
    xmlStream.leafNode('calcPr', {
      calcId: 171027,
      fullCalcOnLoad: model!.fullCalcOnLoad ? 1 : undefined,
    });
  }

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
