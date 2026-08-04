import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        theme: attributes.name ? attributes.name : null,
        showFirstColumn: attributes.showFirstColumn === '1',
        showLastColumn: attributes.showLastColumn === '1',
        showRowStripes: attributes.showRowStripes === '1',
        showColumnStripes: attributes.showColumnStripes === '1',
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

export default TableStyleInfoXform;
export {TableStyleInfoXform};
