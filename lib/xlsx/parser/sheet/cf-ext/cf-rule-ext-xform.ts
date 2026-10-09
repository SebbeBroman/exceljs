import type {CfRuleExtModel} from '../../../xform/sheet/cf-ext/cf-rule-ext-xform.js';
export type {CfRuleExtModel} from '../../../xform/sheet/cf-ext/cf-rule-ext-xform.js';
import BaseXform from '../../../base-parser.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import DatabarExtXform from './databar-ext-xform.js';
import IconSetExtXform from './icon-set-ext-xform.js';

const extIcons: Record<string, boolean> = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

class CfRuleExtXform extends CompositeXform<CfRuleExtModel> {
  databarXform: DatabarExtXform;
  iconSetXform: IconSetExtXform;

  constructor() {
    super();

    this.map = {
      'x14:dataBar': (this.databarXform = new DatabarExtXform()),
      'x14:iconSet': (this.iconSetXform = new IconSetExtXform()),
    };
  }

  override tag = 'x14:cfRule';

  static isExt(rule: CfRuleExtModel): boolean {
    // is this rule primitive?
    if (rule.type === 'dataBar') {
      return DatabarExtXform.isExt(rule);
    }
    if (rule.type === 'iconSet') {
      if (rule.custom || (rule.iconSet && extIcons[rule.iconSet])) {
        return true;
      }
    }
    return false;
  }

  override createNewModel(node?: XmlNode): CfRuleExtModel {
    const attributes = node?.attributes || {};
    return {
      type: attributes.type,
      x14Id: attributes.id,
      priority: BaseXform.toIntValue(attributes.priority),
    };
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    Object.assign(this.model as CfRuleExtModel, parser.model);
  }
}

export default CfRuleExtXform;
export {CfRuleExtXform};
