import type {CfIconExtModel} from '../../../xform/sheet/cf-ext/cf-icon-ext-xform.js';
export type {CfIconExtModel} from '../../../xform/sheet/cf-ext/cf-icon-ext-xform.js';
import BaseXform from '../../../base-parser.js';
import type {XmlNode} from '../../../base-parser.js';

class CfIconExtXform extends BaseXform<CfIconExtModel> {
  override tag = 'x14:cfIcon';

  override parseOpen(node: XmlNode): void {
    const {attributes} = node;
    this.model = {
      iconSet: attributes.iconSet,
      iconId: BaseXform.toIntValue(attributes.iconId),
    };
  }

  override parseClose(name?: string): boolean {
    return name !== this.tag;
  }
}

export default CfIconExtXform;
export {CfIconExtXform};
