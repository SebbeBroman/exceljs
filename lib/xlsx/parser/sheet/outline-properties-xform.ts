import type {OutlinePropertiesModel} from '../../xform/sheet/outline-properties-xform.js';
export type {OutlinePropertiesModel} from '../../xform/sheet/outline-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

const isDefined = (attr: unknown): boolean => typeof attr !== 'undefined';

class OutlinePropertiesXform extends BaseXform<OutlinePropertiesModel> {
  override tag = 'outlinePr';

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
