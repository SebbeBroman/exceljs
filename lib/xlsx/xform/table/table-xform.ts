import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';
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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const {name, attributes} = node;
    switch (name) {
      case this.tag:
        this.reset();
        this.model = {
          name: attributes.name,
          displayName: attributes.displayName || attributes.name,
          tableRef: attributes.ref,
          totalsRow: attributes.totalsRowCount === '1',
          headerRow: attributes.headerRowCount === '1',
        };
        break;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.columns = this.map.tableColumns.model as TableColumnModel[] | undefined;
        if (this.map.autoFilter.model) {
          this.model!.autoFilterRef = this.map.autoFilter.model.autoFilterRef;
          this.map.autoFilter.model.columns.forEach((column, index) => {
            this.model!.columns![index].filterButton = column.filterButton;
          });
        }
        this.model!.style = this.map.tableStyleInfo.model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  override reconcile(model?: TableModel | null, options?: XformOptions): void {
    // fetch the dfxs from styles
    model!.columns!.forEach(column => {
      if (column.dxfId !== undefined) {
        column.style = (options as {styles: {getDxfStyle(id: string | number): unknown}}).styles.getDxfStyle(
          column.dxfId,
        );
      }
    });
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
