import BaseXform from '../../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';

export interface CfIconExtModel {
  iconSet?: string;
  iconId?: number;
}

class CfIconExtXform extends BaseXform<CfIconExtModel> {
  override tag = 'x14:cfIcon';

  override render(xmlStream: XmlStreamLike, model?: CfIconExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.leafNode(this.tag, {
      iconSet: model.iconSet,
      iconId: model.iconId,
    });
  }

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
