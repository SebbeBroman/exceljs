import {v4 as uuidv4} from 'uuid';
import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../../base-xform.js';
import DatabarExtXform from './databar-ext-xform.js';
import IconSetExtXform from './icon-set-ext-xform.js';

const extIcons: Record<string, boolean> = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

export interface CfRuleExtModel {
  type?: string;
  x14Id?: string;
  priority?: number;
  custom?: boolean;
  iconSet?: string;
  gradient?: boolean;
  [key: string]: unknown;
}

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

  override prepare(model?: CfRuleExtModel | null, _options?: XformOptions): void {
    if (model && CfRuleExtXform.isExt(model)) {
      model.x14Id = `{${uuidv4()}}`.toUpperCase();
    }
  }

  override render(xmlStream: XmlStreamLike, model?: CfRuleExtModel | null): void {
    if (!model || !CfRuleExtXform.isExt(model)) {
      return;
    }

    switch (model.type) {
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
    }
  }

  renderDataBar(xmlStream: XmlStreamLike, model: CfRuleExtModel): void {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      id: model.x14Id,
    });

    this.databarXform.render(xmlStream, model as never);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream: XmlStreamLike, model: CfRuleExtModel): void {
    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
      id: model.x14Id || `{${uuidv4()}}`,
    });

    this.iconSetXform.render(xmlStream, model as never);

    xmlStream.closeNode();
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
