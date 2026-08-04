import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';
import CfvoExtXform from './cfvo-ext-xform.js';
import type {CfvoExtModel} from './cfvo-ext-xform.js';
import CfIconExtXform from './cf-icon-ext-xform.js';
import type {CfIconExtModel} from './cf-icon-ext-xform.js';

export interface IconSetExtModel {
  cfvo: CfvoExtModel[];
  iconSet?: string;
  reverse?: boolean;
  showValue?: boolean;
  icons?: CfIconExtModel[];
  custom?: boolean;
  [key: string]: unknown;
}

class IconSetExtXform extends CompositeXform<IconSetExtModel> {
  cfvoXform: CfvoExtXform;
  cfIconXform: CfIconExtXform;

  constructor() {
    super();

    this.map = {
      'x14:cfvo': (this.cfvoXform = new CfvoExtXform()),
      'x14:cfIcon': (this.cfIconXform = new CfIconExtXform()),
    };
  }

  override tag = 'x14:iconSet';

  override render(xmlStream: XmlStreamLike, model?: IconSetExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
      custom: BaseXform.toBoolAttribute(model.icons, false),
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    if (model.icons) {
      model.icons.forEach((icon, i) => {
        icon.iconId = i;
        this.cfIconXform.render(xmlStream, icon);
      });
    }

    xmlStream.closeNode();
  }

  override createNewModel(node?: XmlNode): IconSetExtModel {
    const attributes = node?.attributes || {};
    return {
      cfvo: [],
      iconSet: BaseXform.toStringValue(attributes.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolValue(attributes.reverse, false),
      showValue: BaseXform.toBoolValue(attributes.showValue, true),
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    const [, prop] = name.split(':');
    const m = this.model as IconSetExtModel;
    switch (prop) {
      case 'cfvo':
        m.cfvo.push(parser.model as CfvoExtModel);
        break;

      case 'cfIcon':
        if (!m.icons) {
          m.icons = [];
        }
        m.icons.push(parser.model as CfIconExtModel);
        break;

      default:
        m[prop] = parser.model;
        break;
    }
  }
}

export default IconSetExtXform;
export {IconSetExtXform};
