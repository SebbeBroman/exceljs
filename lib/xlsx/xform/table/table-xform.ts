import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import ListXform from '../list-xform.js';
import AutoFilterXform from './auto-filter-xform.js';
import TableColumnXform from './table-column-xform.js';
import type {TableColumnModel} from './table-column-xform.js';
import TableStyleInfoXform from './table-style-info-xform.js';
import type {TableStyleInfoModel} from './table-style-info-xform.js';

export interface TableModel {
  id?: number | string;
  name?: string;
  displayName?: string;
  tableRef?: string;
  totalsRow?: boolean;
  headerRow?: boolean;
  columns?: TableColumnModel[];
  autoFilterRef?: string;
  style?: TableStyleInfoModel | null;
}

class TableXform extends BaseXform<TableModel> {
  declare map: {
    autoFilter: AutoFilterXform;
    tableColumns: ListXform;
    tableStyleInfo: TableStyleInfoXform;
  };

  constructor() {
    super();

    this.map = {
      autoFilter: new AutoFilterXform(),
      tableColumns: new ListXform({
        tag: 'tableColumns',
        count: true,
        empty: true,
        childXform: new TableColumnXform(),
      }),
      tableStyleInfo: new TableStyleInfoXform(),
    };
  }

  override prepare(model?: TableModel | null, options?: XformOptions): void {
    this.map.autoFilter.prepare(model as never);
    this.map.tableColumns.prepare(model!.columns, options);
  }

  override tag = 'table';

  override render(xmlStream: XmlStreamLike, model?: TableModel | null): void {
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...TableXform.TABLE_ATTRIBUTES,
      id: model!.id,
      name: model!.name,
      displayName: model!.displayName || model!.name,
      ref: model!.tableRef,
      totalsRowCount: model!.totalsRow ? '1' : undefined,
      totalsRowShown: model!.totalsRow ? undefined : '1',
      headerRowCount: model!.headerRow ? '1' : '0',
    });

    this.map.autoFilter.render(xmlStream, model as never);
    this.map.tableColumns.render(xmlStream, model!.columns);
    this.map.tableStyleInfo.render(xmlStream, model!.style);

    xmlStream.closeNode();
  }

  static TABLE_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'xr xr3',
    'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xmlns:xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    // 'xr:uid': '{00000000-000C-0000-FFFF-FFFF00000000}',
  };
}

export default TableXform;
export {TableXform};
