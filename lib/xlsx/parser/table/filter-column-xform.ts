import type {FilterColumnModel} from '../../xform/table/filter-column-xform.js';
export type {FilterColumnModel} from '../../xform/table/filter-column-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import ListXform from '../list-xform.js';
import CustomFilterXform from './custom-filter-xform.js';
import FilterXform from './filter-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    const {attributes} = node;
    switch (node.name) {
      case this.tag:
        this.model = {
          filterButton: attributes.hiddenButton === '0',
        };
        return true;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parseOpen(node);
          return true;
        }
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.customFilters = this.map.customFilters.model as unknown[] | undefined;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }
}

export default FilterColumnXform;
export {FilterColumnXform};
