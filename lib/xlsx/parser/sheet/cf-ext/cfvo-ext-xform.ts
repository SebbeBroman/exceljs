import type {CfvoExtModel} from '../../../xform/sheet/cf-ext/cfvo-ext-xform.js';
export type {CfvoExtModel} from '../../../xform/sheet/cf-ext/cfvo-ext-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import FExtXform from './f-ext-xform.js';

class CfvoExtXform extends CompositeXform<CfvoExtModel> {
  fExtXform: FExtXform;

  constructor() {
    super();

    this.map = {
      'xm:f': (this.fExtXform = new FExtXform()),
    };
  }

  override tag = 'x14:cfvo';

  override createNewModel(node?: XmlNode): CfvoExtModel {
    return {
      type: node?.attributes.type,
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    switch (name) {
      case 'xm:f':
        (this.model as CfvoExtModel).value = parser.model ? parseFloat(parser.model as string) : 0;
        break;
    }
  }
}

export default CfvoExtXform;
export {CfvoExtXform};
