import type {DatabarModel} from '../../../xform/sheet/cf/databar-xform.js';
export type {DatabarModel} from '../../../xform/sheet/cf/databar-xform.js';
import CompositeXform from '../../composite-xform.js';
import ColorXform from '../../style/color-xform.js';
import type {ColorModel} from '../../style/color-xform.js';
import CfvoXform from './cfvo-xform.js';
import type {CfvoModel} from './cfvo-xform.js';

class DatabarXform extends CompositeXform<DatabarModel> {
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform()),
    };
  }

  override tag = 'dataBar';

  override createNewModel(): DatabarModel {
    return {
      cfvo: [],
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    switch (name) {
      case 'cfvo':
        (this.model as DatabarModel).cfvo.push(parser.model as CfvoModel);
        break;
      case 'color':
        (this.model as DatabarModel).color = parser.model as ColorModel;
        break;
    }
  }
}

export default DatabarXform;
export {DatabarXform};
