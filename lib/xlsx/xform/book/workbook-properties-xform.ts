import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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
