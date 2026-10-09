import type {ColorScaleModel} from '../../../xform/sheet/cf/color-scale-xform.js';
export type {ColorScaleModel} from '../../../xform/sheet/cf/color-scale-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import ColorXform from '../../style/color-xform.js';
import CfvoXform from './cfvo-xform.js';

class ColorScaleXform extends CompositeXform<ColorScaleModel> {
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform()),
    };
  }

  override tag = 'colorScale';

  override createNewModel(_node?: XmlNode): ColorScaleModel {
    return {
      cfvo: [],
      color: [],
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    (this.model as ColorScaleModel)[name as 'cfvo' | 'color'].push(parser.model as never);
  }
}

export default ColorScaleXform;
export {ColorScaleXform};
