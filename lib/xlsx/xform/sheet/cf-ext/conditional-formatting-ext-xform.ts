import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XformOptions} from '../../base-xform.js';
import SqRefExtXform from './sqref-ext-xform.js';
import CfRuleExtXform from './cf-rule-ext-xform.js';
import type {CfRuleExtModel} from './cf-rule-ext-xform.js';

export interface ConditionalFormattingExtModel {
  ref?: string;
  rules: CfRuleExtModel[];
}

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

  override prepare(
    model?: ConditionalFormattingExtModel | null,
    options?: XformOptions,
  ): void {
    if (!model) {
      return;
    }
    model.rules.forEach(rule => {
      this.cfRule.prepare(rule, options);
    });
  }

  override render(
    xmlStream: XmlStreamLike,
    model?: ConditionalFormattingExtModel | null,
  ): void {
    if (!model || !model.rules.some(CfRuleExtXform.isExt)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      'xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main',
    });

    model.rules.filter(CfRuleExtXform.isExt).forEach(rule => this.cfRule.render(xmlStream, rule));

    // for some odd reason, Excel needs the <xm:sqref> node to be after the rules
    this.sqRef.render(xmlStream, model.ref);

    xmlStream.closeNode();
  }

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
