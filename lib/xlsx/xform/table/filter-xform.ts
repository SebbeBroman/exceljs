import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface FilterModel {
  val?: string;
}

class FilterXform extends BaseXform<FilterModel> {
  override tag = 'filter';

  override render(xmlStream: XmlStreamLike, model?: FilterModel | null): void {
    xmlStream.leafNode(this.tag, {
      val: model!.val,
    });
  }
}

export default FilterXform;
export {FilterXform};
