import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

export interface VmlPositionXformOptions {
  tag: string;
}

class VmlPositionXform extends BaseXform<Record<string, boolean>> {
  _model: VmlPositionXformOptions;
  override tag: string;

  constructor(model: VmlPositionXformOptions) {
    super();
    this._model = model;
    this.tag = model.tag;
  }

  override render(xmlStream: XmlStreamLike, model?: unknown, index?: number): void {
    // Third argument is position-type list at call sites (not a numeric index).
    const type = index as unknown as string[] | undefined;
    if (model === type![2]) {
      xmlStream.leafNode(this.tag);
    } else if (this.tag === 'x:SizeWithCells' && model === type![1]) {
      xmlStream.leafNode(this.tag);
    }
  }
}

export default VmlPositionXform;
export {VmlPositionXform};
