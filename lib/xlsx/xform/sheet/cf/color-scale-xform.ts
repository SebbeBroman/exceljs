import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';
import ColorXform from '../../style/color-xform.js';
import type {ColorModel} from '../../style/color-xform.js';
import CfvoXform from './cfvo-xform.js';
import type {CfvoModel} from './cfvo-xform.js';

export interface ColorScaleModel {
  cfvo: CfvoModel[];
  color: ColorModel[];
}

class ColorScaleXform extends CompositeXform<ColorScaleModel> {
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform()),
    };
  }

  override tag = 'colorScale';

  override render(xmlStream: XmlStreamLike, model?: ColorScaleModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach(color => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  override createNewModel(_node?: XmlNode): ColorScaleModel {
    return {
      cfvo: [],
      color: [],
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    (this.model as ColorScaleModel)[name as 'cfvo' | 'color'].push(parser.model as never);
  }
}

export default ColorScaleXform;
export {ColorScaleXform};
