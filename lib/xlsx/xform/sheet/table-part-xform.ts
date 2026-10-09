import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface TablePartModel {
  rId: string;
}

class TablePartXform extends BaseXform<TablePartModel> {
  override tag = 'tablePart';

  override render(xmlStream: XmlStreamLike, model?: TablePartModel | null): void {
    if (model) {
      xmlStream.leafNode(this.tag, {'r:id': model.rId});
    }
  }
}

export default TablePartXform;
export {TablePartXform};
