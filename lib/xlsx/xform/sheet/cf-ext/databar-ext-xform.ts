import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';
import ColorXform from '../../style/color-xform.js';
import type {ColorModel} from '../../style/color-xform.js';
import CfvoExtXform from './cfvo-ext-xform.js';
import type {CfvoExtModel} from './cfvo-ext-xform.js';

export interface DatabarExtModel {
  cfvo: CfvoExtModel[];
  minLength?: number;
  maxLength?: number;
  border?: boolean;
  gradient?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: string;
  direction?: string;
  borderColor?: ColorModel;
  negativeBorderColor?: ColorModel;
  negativeFillColor?: ColorModel;
  axisColor?: ColorModel;
  [key: string]: unknown;
}

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

  override render(xmlStream: XmlStreamLike, model?: DatabarExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag, {
      minLength: BaseXform.toIntAttribute(model.minLength, 0, true),
      maxLength: BaseXform.toIntAttribute(model.maxLength, 100, true),
      border: BaseXform.toBoolAttribute(model.border, false),
      gradient: BaseXform.toBoolAttribute(model.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarColorSameAsPositive,
        true,
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarBorderColorSameAsPositive,
        true,
      ),
      axisPosition: BaseXform.toAttribute(model.axisPosition, 'auto'),
      direction: BaseXform.toAttribute(model.direction, 'leftToRight'),
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    this.borderColorXform.render(xmlStream, model.borderColor);
    this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor);
    this.negativeFillColorXform.render(xmlStream, model.negativeFillColor);
    this.axisColorXform.render(xmlStream, model.axisColor);

    xmlStream.closeNode();
  }

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
