import type {TableModel} from '../../xform/table/table-xform.js';
export type {TableModel} from '../../xform/table/table-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode, XformOptions} from '../../base-parser.js';
import ListXform from '../list-xform.js';
import AutoFilterXform from './auto-filter-xform.js';
import TableColumnXform from './table-column-xform.js';
import type {TableColumnModel} from './table-column-xform.js';
import TableStyleInfoXform from './table-style-info-xform.js';

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

  override tag = 'table';

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
        column.style = (
          options as {styles: {getDxfStyle(id: string | number): unknown}}
        ).styles.getDxfStyle(column.dxfId);
      }
    });
  }
}

export default TableXform;
export {TableXform};
