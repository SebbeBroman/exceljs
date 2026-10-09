import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

class DimensionXform extends BaseXform<string> {
  override tag = 'dimension';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    if (model) {
      xmlStream.leafNode('dimension', {ref: model});
    }
  }
}

export default DimensionXform;
export {DimensionXform};
