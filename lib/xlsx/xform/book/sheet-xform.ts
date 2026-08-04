import utils from '../../../utils/utils.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'sheet') {
      this.model = {
        name: utils.xmlDecode(node.attributes.name),
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
