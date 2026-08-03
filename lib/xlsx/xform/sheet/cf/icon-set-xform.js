import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import CfvoXform from './cfvo-xform.js';

class IconSetXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
    };
  }

  get tag() {
    return 'iconSet';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    xmlStream.closeNode();
  }

  createNewModel({attributes}) {
    return {
      iconSet: BaseXform.toStringValue(attributes.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolValue(attributes.reverse),
      showValue: BaseXform.toBoolValue(attributes.showValue),
      cfvo: [],
    };
  }

  onParserClose(name, parser) {
    this.model[name].push(parser.model);
  }
}

export default IconSetXform;
export { IconSetXform };
