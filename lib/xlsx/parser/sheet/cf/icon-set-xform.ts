import type {IconSetModel} from '../../../xform/sheet/cf/icon-set-xform.js';
export type {IconSetModel} from '../../../xform/sheet/cf/icon-set-xform.js';
import BaseXform from '../../../base-parser.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import CfvoXform from './cfvo-xform.js';
import type {CfvoModel} from './cfvo-xform.js';

class IconSetXform extends CompositeXform<IconSetModel> {
  cfvoXform: CfvoXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
    };
  }

  override tag = 'iconSet';

  override createNewModel(node?: XmlNode): IconSetModel {
    const attributes = node?.attributes || {};
    return {
      iconSet: BaseXform.toStringValue(attributes.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolValue(attributes.reverse),
      showValue: BaseXform.toBoolValue(attributes.showValue),
      cfvo: [],
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    (this.model as IconSetModel).cfvo.push(parser.model as CfvoModel);
    void name;
  }
}

export default IconSetXform;
export {IconSetXform};
