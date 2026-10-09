import type {ConditionalFormattingModel} from '../../../xform/sheet/cf/conditional-formatting-xform.js';
export type {ConditionalFormattingModel} from '../../../xform/sheet/cf/conditional-formatting-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import CfRuleXform from './cf-rule-xform.js';
import type {CfRuleModel} from './cf-rule-xform.js';

class ConditionalFormattingXform extends CompositeXform<ConditionalFormattingModel> {
  constructor() {
    super();

    this.map = {
      cfRule: new CfRuleXform(),
    };
  }

  override tag = 'conditionalFormatting';

  override createNewModel(node?: XmlNode): ConditionalFormattingModel {
    return {
      ref: node?.attributes.sqref as string,
      rules: [],
    };
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    (this.model as ConditionalFormattingModel).rules.push(parser.model as CfRuleModel);
  }
}

export default ConditionalFormattingXform;
export {ConditionalFormattingXform};
