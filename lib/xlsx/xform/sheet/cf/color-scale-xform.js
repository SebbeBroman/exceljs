import CompositeXform from '../../composite-xform.js';
import ColorXform from '../../style/color-xform.js';
import CfvoXform from './cfvo-xform.js';

class ColorScaleXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform()),
    };
  }

  get tag() {
    return 'colorScale';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach(color => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  createNewModel(node) {
    return {
      cfvo: [],
      color: [],
    };
  }

  onParserClose(name, parser) {
    this.model[name].push(parser.model);
  }
}

export default ColorScaleXform;
export { ColorScaleXform };
