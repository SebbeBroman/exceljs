import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

class FormulaXform extends BaseXform<string> {
  override tag = 'formula';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }
}

export default FormulaXform;
export {FormulaXform};
