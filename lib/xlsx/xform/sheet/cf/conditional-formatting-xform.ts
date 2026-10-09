import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';
import CfRuleXform from './cf-rule-xform.js';
import type {CfRuleModel} from './cf-rule-xform.js';

export interface ConditionalFormattingModel {
  ref: string;
  rules: CfRuleModel[];
}

class ConditionalFormattingXform extends CompositeXform<ConditionalFormattingModel> {
  constructor() {
    super();

    this.map = {
      cfRule: new CfRuleXform(),
    };
  }

  override tag = 'conditionalFormatting';

  override render(xmlStream: XmlStreamLike, model?: ConditionalFormattingModel | null): void {
    if (!model) {
      return;
    }
    // if there are no primitive rules, exit now
    if (!model.rules.some(CfRuleXform.isPrimitive)) {
      return;
    }

    xmlStream.openNode(this.tag, {sqref: model.ref});

    model.rules.forEach(rule => {
      if (CfRuleXform.isPrimitive(rule)) {
        rule.ref = model.ref;
        this.map.cfRule.render(xmlStream, rule);
      }
    });

    xmlStream.closeNode();
  }
}

export default ConditionalFormattingXform;
export {ConditionalFormattingXform};
