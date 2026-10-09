import type {ConditionalFormattingsOptions} from '../../../xform/sheet/cf/conditional-formattings-xform.js';
export type {
  ConditionalFormattingsOptions,
  StylesDxfLike,
} from '../../../xform/sheet/cf/conditional-formattings-xform.js';
import BaseXform from '../../../base-parser.js';
import type {XmlNode} from '../../../base-parser.js';
import ConditionalFormattingXform from './conditional-formatting-xform.js';
import type {ConditionalFormattingModel} from './conditional-formatting-xform.js';
import type {CfRuleModel} from './cf-rule-xform.js';

class ConditionalFormattingsXform extends BaseXform<ConditionalFormattingModel[]> {
  cfXform: ConditionalFormattingXform;

  constructor() {
    super();

    this.cfXform = new ConditionalFormattingXform();
  }

  override tag = 'conditionalFormatting';

  override reset(): void {
    this.model = [];
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case 'conditionalFormatting':
        this.parser = this.cfXform;
        this.parser.parseOpen(node);
        return true;

      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        (this.model as ConditionalFormattingModel[]).push(
          this.parser.model as ConditionalFormattingModel,
        );
        this.parser = undefined;
        return false;
      }
      return true;
    }
    return false;
  }

  override reconcile(
    model?: ConditionalFormattingModel[] | null,
    options?: ConditionalFormattingsOptions,
  ): void {
    if (!model || !options) {
      return;
    }
    model.forEach(cf => {
      cf.rules.forEach((rule: CfRuleModel) => {
        if (rule.dxfId !== undefined) {
          rule.style = options.styles.getDxfStyle(rule.dxfId);
          delete rule.dxfId;
        }
      });
    });
  }
}

export default ConditionalFormattingsXform;
export {ConditionalFormattingsXform};
