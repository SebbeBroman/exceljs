import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import ListXform from '../list-xform.js';
import CustomFilterXform from './custom-filter-xform.js';
import FilterXform from './filter-xform.js';

export interface FilterColumnModel {
  colId?: string;
  filterButton?: boolean;
  customFilters?: unknown[];
}

class FilterColumnXform extends BaseXform<FilterColumnModel> {
  declare map: {
    customFilters: ListXform;
    filters: ListXform;
  };

  constructor() {
    super();

    this.map = {
      customFilters: new ListXform({
        tag: 'customFilters',
        count: false,
        empty: true,
        childXform: new CustomFilterXform(),
      }),
      filters: new ListXform({
        tag: 'filters',
        count: false,
        empty: true,
        childXform: new FilterXform(),
      }),
    };
  }

  override tag = 'filterColumn';

  override prepare(model?: FilterColumnModel | null, options?: XformOptions): void {
    model!.colId = options!.index!.toString();
  }

  override render(xmlStream: XmlStreamLike, model?: FilterColumnModel | null): boolean {
    if (model!.customFilters) {
      xmlStream.openNode(this.tag, {
        colId: model!.colId,
        hiddenButton: model!.filterButton ? '0' : '1',
      });

      this.map.customFilters.render(xmlStream, model!.customFilters);

      xmlStream.closeNode();
      return true;
    }
    xmlStream.leafNode(this.tag, {
      colId: model!.colId,
      hiddenButton: model!.filterButton ? '0' : '1',
    });
    return true;
  }
}

export default FilterColumnXform;
export {FilterColumnXform};
