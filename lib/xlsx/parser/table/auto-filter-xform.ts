import type {AutoFilterModel} from '../../xform/table/auto-filter-xform.js';
export type {AutoFilterModel} from '../../xform/table/auto-filter-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import FilterColumnXform from './filter-column-xform.js';
import type {FilterColumnModel} from './filter-column-xform.js';

class AutoFilterXform extends BaseXform<AutoFilterModel> {
  declare map: {
    filterColumn: FilterColumnXform;
  };

  constructor() {
    super();

    this.map = {
      filterColumn: new FilterColumnXform(),
    };
  }

  override tag = 'autoFilter';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {
          autoFilterRef: node.attributes.ref,
          columns: [],
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

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.columns.push(this.parser.model as FilterColumnModel);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export default AutoFilterXform;
export {AutoFilterXform};
