import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';
import CfvoXform from './cfvo-xform.js';
import type {CfvoModel} from './cfvo-xform.js';

export interface IconSetModel {
  iconSet?: string;
  reverse?: boolean;
  showValue?: boolean;
  cfvo: CfvoModel[];
  custom?: boolean;
}

class IconSetXform extends CompositeXform<IconSetModel> {
  cfvoXform: CfvoXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
    };
  }

  override tag = 'iconSet';

  override render(xmlStream: XmlStreamLike, model?: IconSetModel | null): void {
    if (!model) {
      return;
    }
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
}

export default IconSetXform;
export {IconSetXform};
