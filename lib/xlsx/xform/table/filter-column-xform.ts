import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';
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
