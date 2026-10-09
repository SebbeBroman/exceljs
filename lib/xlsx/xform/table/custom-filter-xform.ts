import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface CustomFilterModel {
  val?: string;
  operator?: string;
}

class CustomFilterXform extends BaseXform<CustomFilterModel> {
  override tag = 'customFilter';

  override render(xmlStream: XmlStreamLike, model?: CustomFilterModel | null): void {
    xmlStream.leafNode(this.tag, {
      val: model!.val,
      operator: model!.operator,
    });
  }
}

export default CustomFilterXform;
export {CustomFilterXform};
