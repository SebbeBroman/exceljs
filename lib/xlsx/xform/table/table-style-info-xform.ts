import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface TableStyleInfoModel {
  theme?: string | null;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

class TableStyleInfoXform extends BaseXform<TableStyleInfoModel> {
  override tag = 'tableStyleInfo';

  override render(xmlStream: XmlStreamLike, model?: TableStyleInfoModel | null): boolean {
    xmlStream.leafNode(this.tag, {
      name: model!.theme ? model!.theme : undefined,
      showFirstColumn: model!.showFirstColumn ? '1' : '0',
      showLastColumn: model!.showLastColumn ? '1' : '0',
      showRowStripes: model!.showRowStripes ? '1' : '0',
      showColumnStripes: model!.showColumnStripes ? '1' : '0',
    });
    return true;
  }
}

export default TableStyleInfoXform;
export {TableStyleInfoXform};
