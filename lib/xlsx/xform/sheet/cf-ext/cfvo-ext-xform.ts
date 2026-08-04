import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';
import FExtXform from './f-ext-xform.js';

export interface CfvoExtModel {
  type?: string;
  value?: number;
}

class CfvoExtXform extends CompositeXform<CfvoExtModel> {
  fExtXform: FExtXform;

  constructor() {
    super();

    this.map = {
      'xm:f': (this.fExtXform = new FExtXform()),
    };
  }

  override tag = 'x14:cfvo';

  override render(xmlStream: XmlStreamLike, model?: CfvoExtModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag, {
      type: model.type,
    });
    if (model.value !== undefined) {
      this.fExtXform.render(xmlStream, model.value);
    }
    xmlStream.closeNode();
  }

  override createNewModel(node?: XmlNode): CfvoExtModel {
    return {
      type: node?.attributes.type,
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    switch (name) {
      case 'xm:f':
        (this.model as CfvoExtModel).value = parser.model ? parseFloat(parser.model as string) : 0;
        break;
    }
  }
}

export default CfvoExtXform;
export {CfvoExtXform};
