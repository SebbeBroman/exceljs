import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';
import ColorXform from '../../style/color-xform.js';
import type {ColorModel} from '../../style/color-xform.js';
import CfvoXform from './cfvo-xform.js';
import type {CfvoModel} from './cfvo-xform.js';

export interface DatabarModel {
  cfvo: CfvoModel[];
  color?: ColorModel;
}

class DatabarXform extends CompositeXform<DatabarModel> {
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform()),
    };
  }

  override tag = 'dataBar';

  override render(xmlStream: XmlStreamLike, model?: DatabarModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    this.colorXform.render(xmlStream, model.color);

    xmlStream.closeNode();
  }
}

export default DatabarXform;
export {DatabarXform};
