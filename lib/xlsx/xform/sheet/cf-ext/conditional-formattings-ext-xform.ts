import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XformOptions} from '../../base-xform.js';
import CfRuleExtXform from './cf-rule-ext-xform.js';
import ConditionalFormattingExtXform from './conditional-formatting-ext-xform.js';
import type {ConditionalFormattingExtModel} from './conditional-formatting-ext-xform.js';

export type ConditionalFormattingsExtModel = ConditionalFormattingExtModel[] & {
  hasExtContent?: boolean;
};

class ConditionalFormattingsExtXform extends CompositeXform<ConditionalFormattingsExtModel> {
  cfXform: ConditionalFormattingExtXform;

  constructor() {
    super();

    this.map = {
      'x14:conditionalFormatting': (this.cfXform = new ConditionalFormattingExtXform()),
    };
  }

  override tag = 'x14:conditionalFormattings';

  hasContent(model?: ConditionalFormattingsExtModel | null): boolean {
    if (!model) {
      return false;
    }
    if (model.hasExtContent === undefined) {
      model.hasExtContent = model.some(cf => cf.rules.some(CfRuleExtXform.isExt));
    }
    return model.hasExtContent;
  }

  override prepare(model?: ConditionalFormattingsExtModel | null, options?: XformOptions): void {
    if (!model) {
      return;
    }
    model.forEach(cf => {
      this.cfXform.prepare(cf, options);
    });
  }

  override render(xmlStream: XmlStreamLike, model?: ConditionalFormattingsExtModel | null): void {
    if (model && this.hasContent(model)) {
      xmlStream.openNode(this.tag);
      model.forEach(cf => this.cfXform.render(xmlStream, cf));
      xmlStream.closeNode();
    }
  }
}

export default ConditionalFormattingsExtXform;
export {ConditionalFormattingsExtXform};
