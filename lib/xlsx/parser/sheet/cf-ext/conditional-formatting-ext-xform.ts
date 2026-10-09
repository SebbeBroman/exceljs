import type {ConditionalFormattingExtModel} from '../../../xform/sheet/cf-ext/conditional-formatting-ext-xform.js';
export type {ConditionalFormattingExtModel} from '../../../xform/sheet/cf-ext/conditional-formatting-ext-xform.js';
import CompositeXform from '../../composite-xform.js';
import SqRefExtXform from './sqref-ext-xform.js';
import CfRuleExtXform from './cf-rule-ext-xform.js';
import type {CfRuleExtModel} from './cf-rule-ext-xform.js';

class ConditionalFormattingExtXform extends CompositeXform<ConditionalFormattingExtModel> {
  sqRef: SqRefExtXform;
  cfRule: CfRuleExtXform;

  constructor() {
    super();

    this.map = {
      'xm:sqref': (this.sqRef = new SqRefExtXform()),
      'x14:cfRule': (this.cfRule = new CfRuleExtXform()),
    };
  }

  override tag = 'x14:conditionalFormatting';

  override createNewModel(): ConditionalFormattingExtModel {
    return {
      rules: [],
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    switch (name) {
      case 'xm:sqref':
        (this.model as ConditionalFormattingExtModel).ref = parser.model as string;
        break;

      case 'x14:cfRule':
        (this.model as ConditionalFormattingExtModel).rules.push(parser.model as CfRuleExtModel);
        break;
    }
  }
}

export default ConditionalFormattingExtXform;
export {ConditionalFormattingExtXform};
