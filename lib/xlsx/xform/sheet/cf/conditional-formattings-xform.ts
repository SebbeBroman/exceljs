import BaseXform from '../../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../../base-xform.js';
import ConditionalFormattingXform from './conditional-formatting-xform.js';
import type {ConditionalFormattingModel} from './conditional-formatting-xform.js';
import type {CfRuleModel} from './cf-rule-xform.js';

export interface StylesDxfLike {
  addDxfStyle(style: unknown): number;
  getDxfStyle(dxfId: number): unknown;
}

export interface ConditionalFormattingsOptions extends XformOptions {
  styles: StylesDxfLike;
}

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

  override prepare(
    model?: ConditionalFormattingModel[] | null,
    options?: ConditionalFormattingsOptions,
  ): void {
    if (!model || !options) {
      return;
    }
    // ensure each rule has a priority value
    let nextPriority = model.reduce(
      (p, cf) => Math.max(p, ...cf.rules.map(rule => rule.priority || 0)),
      1,
    );
    model.forEach(cf => {
      cf.rules.forEach(rule => {
        if (!rule.priority) {
          rule.priority = nextPriority++;
        }

        if (rule.style) {
          rule.dxfId = options.styles.addDxfStyle(rule.style);
        }
      });
    });
  }

  override render(
    xmlStream: XmlStreamLike,
    model?: ConditionalFormattingModel[] | null,
  ): void {
    if (!model) {
      return;
    }
    model.forEach(cf => {
      this.cfXform.render(xmlStream, cf);
    });
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
