import type {CfvoModel} from '../../../xform/sheet/cf/cfvo-xform.js';
export type {CfvoModel} from '../../../xform/sheet/cf/cfvo-xform.js';
import BaseXform from '../../../base-parser.js';
import type {XmlNode} from '../../../base-parser.js';

class CfvoXform extends BaseXform<CfvoModel> {
  override tag = 'cfvo';

  override parseOpen(node: XmlNode): void {
    this.model = {
      type: node.attributes.type,
      value: BaseXform.toFloatValue(node.attributes.val),
    };
  }

  override parseClose(name?: string): boolean {
    return name !== this.tag;
  }
}

export default CfvoXform;
export {CfvoXform};
