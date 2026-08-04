import BaseXform from '../../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';

export interface CfvoModel {
  type: string;
  value?: number;
}

class CfvoXform extends BaseXform<CfvoModel> {
  override tag = 'cfvo';

  override render(xmlStream: XmlStreamLike, model?: CfvoModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.leafNode(this.tag, {
      type: model.type,
      val: model.value,
    });
  }

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
