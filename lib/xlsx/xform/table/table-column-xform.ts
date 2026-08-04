import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';

export interface TableColumnModel {
  id?: number;
  name?: string;
  totalsRowLabel?: string;
  totalsRowFunction?: string;
  dxfId?: string | number;
  filterButton?: boolean;
  style?: unknown;
}

class TableColumnXform extends BaseXform<TableColumnModel> {
  override tag = 'tableColumn';

  override prepare(model?: TableColumnModel | null, options?: XformOptions): void {
    model!.id = (options!.index as number) + 1;
  }

  override render(xmlStream: XmlStreamLike, model?: TableColumnModel | null): boolean {
    xmlStream.leafNode(this.tag, {
      id: model!.id!.toString(),
      name: model!.name,
      totalsRowLabel: model!.totalsRowLabel,
      totalsRowFunction: model!.totalsRowFunction,
      dxfId: model!.dxfId,
    });
    return true;
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      const {attributes} = node;
      this.model = {
        name: attributes.name,
        totalsRowLabel: attributes.totalsRowLabel,
        totalsRowFunction: attributes.totalsRowFunction,
        dxfId: attributes.dxfId,
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

export default TableColumnXform;
export {TableColumnXform};
