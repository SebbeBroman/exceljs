import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface SheetModel {
  id: number;
  name: string;
  state?: string;
  rId: string;
}

class WorksheetXform extends BaseXform<SheetModel> {
  override render(xmlStream: XmlStreamLike, model?: SheetModel | null): void {
    xmlStream.leafNode('sheet', {
      sheetId: model!.id,
      name: model!.name,
      state: model!.state,
      'r:id': model!.rId,
    });
  }
}

export default WorksheetXform;
export {WorksheetXform};
