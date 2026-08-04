import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface OutlinePropertiesModel {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}

const isDefined = (attr: unknown): boolean => typeof attr !== 'undefined';

class OutlinePropertiesXform extends BaseXform<OutlinePropertiesModel> {
  override tag = 'outlinePr';

  override render(xmlStream: XmlStreamLike, model?: OutlinePropertiesModel | null): boolean {
    if (model && (isDefined(model.summaryBelow) || isDefined(model.summaryRight))) {
      xmlStream.leafNode(this.tag, {
        summaryBelow: isDefined(model.summaryBelow) ? Number(model.summaryBelow) : undefined,
        summaryRight: isDefined(model.summaryRight) ? Number(model.summaryRight) : undefined,
      });
      return true;
    }
    return false;
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      this.model = {
        summaryBelow: isDefined(node.attributes.summaryBelow)
          ? Boolean(Number(node.attributes.summaryBelow))
          : undefined,
        summaryRight: isDefined(node.attributes.summaryRight)
          ? Boolean(Number(node.attributes.summaryRight))
          : undefined,
      };
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default OutlinePropertiesXform;
export {OutlinePropertiesXform};
