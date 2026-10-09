import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import FilterColumnXform from './filter-column-xform.js';
import type {FilterColumnModel} from './filter-column-xform.js';

export interface AutoFilterModel {
  autoFilterRef?: string;
  columns: FilterColumnModel[];
}

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

  override prepare(model?: AutoFilterModel | null): void {
    model!.columns.forEach((column, index) => {
      this.map.filterColumn.prepare(column, {index});
    });
  }

  override render(xmlStream: XmlStreamLike, model?: AutoFilterModel | null): boolean {
    xmlStream.openNode(this.tag, {ref: model!.autoFilterRef});

    model!.columns.forEach(column => {
      this.map.filterColumn.render(xmlStream, column);
    });

    xmlStream.closeNode();
    return true;
  }
}

export default AutoFilterXform;
export {AutoFilterXform};
