import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';

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
}

export default TableColumnXform;
export {TableColumnXform};
