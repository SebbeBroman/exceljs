import type {ColorModel} from '../../xform/style/color-xform.js';
export type {ColorModel} from '../../xform/style/color-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class ColorXform extends BaseXform<ColorModel> {
  name: string;

  constructor(name?: string) {
    super();

    // this.name controls the xm node name
    this.name = name || 'color';
    this.tag = this.name;
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.name) {
      if (node.attributes.rgb) {
        this.model = {argb: node.attributes.rgb};
      } else if (node.attributes.theme) {
        this.model = {theme: parseInt(node.attributes.theme, 10)};
        if (node.attributes.tint) {
          this.model.tint = parseFloat(node.attributes.tint);
        }
      } else if (node.attributes.indexed) {
        this.model = {indexed: parseInt(node.attributes.indexed, 10)};
      } else {
        this.model = undefined;
      }
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default ColorXform;
export {ColorXform};
