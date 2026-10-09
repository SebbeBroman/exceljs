import type {DatabarExtModel} from '../../../xform/sheet/cf-ext/databar-ext-xform.js';
export type {DatabarExtModel} from '../../../xform/sheet/cf-ext/databar-ext-xform.js';
import BaseXform from '../../../base-parser.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import ColorXform from '../../style/color-xform.js';
import CfvoExtXform from './cfvo-ext-xform.js';
import type {CfvoExtModel} from './cfvo-ext-xform.js';

class DatabarExtXform extends CompositeXform<DatabarExtModel> {
  cfvoXform: CfvoExtXform;
  borderColorXform: ColorXform;
  negativeBorderColorXform: ColorXform;
  negativeFillColorXform: ColorXform;
  axisColorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      'x14:cfvo': (this.cfvoXform = new CfvoExtXform()),
      'x14:borderColor': (this.borderColorXform = new ColorXform('x14:borderColor')),
      'x14:negativeBorderColor': (this.negativeBorderColorXform = new ColorXform(
        'x14:negativeBorderColor',
      )),
      'x14:negativeFillColor': (this.negativeFillColorXform = new ColorXform(
        'x14:negativeFillColor',
      )),
      'x14:axisColor': (this.axisColorXform = new ColorXform('x14:axisColor')),
    };
  }

  static isExt(rule: {gradient?: boolean}): boolean {
    // not all databars need ext
    // TODO: refine this
    return !rule.gradient;
  }

  override tag = 'x14:dataBar';

  override createNewModel(node?: XmlNode): DatabarExtModel {
    const attributes = node?.attributes || {};
    return {
      cfvo: [],
      minLength: BaseXform.toIntValue(attributes.minLength, 0),
      maxLength: BaseXform.toIntValue(attributes.maxLength, 100),
      border: BaseXform.toBoolValue(attributes.border, false),
      gradient: BaseXform.toBoolValue(attributes.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarColorSameAsPositive,
        true,
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarBorderColorSameAsPositive,
        true,
      ),
      axisPosition: BaseXform.toStringValue(attributes.axisPosition, 'auto'),
      direction: BaseXform.toStringValue(attributes.direction, 'leftToRight'),
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    const [, prop] = name.split(':');
    switch (prop) {
      case 'cfvo':
        (this.model as DatabarExtModel).cfvo.push(parser.model as CfvoExtModel);
        break;

      default:
        (this.model as DatabarExtModel)[prop] = parser.model;
        break;
    }
  }
}

export default DatabarExtXform;
export {DatabarExtXform};
