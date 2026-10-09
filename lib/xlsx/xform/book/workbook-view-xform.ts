import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface WorkbookViewModel {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  firstSheet?: number;
  activeTab?: number;
  visibility?: string;
}

class WorkbookViewXform extends BaseXform<WorkbookViewModel> {
  override render(xmlStream: XmlStreamLike, model?: WorkbookViewModel | null): void {
    const attributes: Record<string, unknown> = {
      xWindow: model!.x || 0,
      yWindow: model!.y || 0,
      windowWidth: model!.width || 12000,
      windowHeight: model!.height || 24000,
      firstSheet: model!.firstSheet,
      activeTab: model!.activeTab,
    };
    if (model!.visibility && model!.visibility !== 'visible') {
      attributes.visibility = model!.visibility;
    }
    xmlStream.leafNode('workbookView', attributes);
  }
}

export default WorkbookViewXform;
export {WorkbookViewXform};
