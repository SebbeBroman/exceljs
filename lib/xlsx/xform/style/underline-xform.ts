import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export type UnderlineModel = boolean | string;

class UnderlineXform extends BaseXform<UnderlineModel> {
  constructor(model?: UnderlineModel) {
    super();

    this.model = model;
  }

  override tag = 'u';

  override render(xmlStream: XmlStreamLike, model?: UnderlineModel | null): void {
    model = model || this.model;

    if (model === true) {
      xmlStream.leafNode('u');
    } else {
      const attr = UnderlineXform.Attributes[model as string];
      if (attr) {
        xmlStream.leafNode('u', attr);
      }
    }
  }

  static Attributes: Record<string, Record<string, string>> = {
    single: {},
    double: {val: 'double'},
    singleAccounting: {val: 'singleAccounting'},
    doubleAccounting: {val: 'doubleAccounting'},
  };
}

export default UnderlineXform;
export {UnderlineXform};
