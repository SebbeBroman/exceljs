import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface WorkbookPropertiesModel {
  date1904?: boolean;
}

class WorksheetPropertiesXform extends BaseXform<WorkbookPropertiesModel> {
  override render(xmlStream: XmlStreamLike, model?: WorkbookPropertiesModel | null): void {
    xmlStream.leafNode('workbookPr', {
      date1904: model!.date1904 ? 1 : undefined,
      defaultThemeVersion: 164011,
      filterPrivacy: 1,
    });
  }
}

export default WorksheetPropertiesXform;
export {WorksheetPropertiesXform};
