import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default WorkbookCalcPropertiesXform;
export {WorkbookCalcPropertiesXform};
