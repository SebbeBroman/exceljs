import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

class MergeCellXform extends BaseXform<string> {
  override tag = 'mergeCell';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.leafNode('mergeCell', {ref: model as string});
  }
}

export default MergeCellXform;
export {MergeCellXform};
