import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

// Color encapsulates translation from color model to/from xlsx
export interface ColorModel {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

class ColorXform extends BaseXform<ColorModel> {
  name: string;

  constructor(name?: string) {
    super();

    // this.name controls the xm node name
    this.name = name || 'color';
    this.tag = this.name;
  }

  override render(xmlStream: XmlStreamLike, model?: ColorModel | null): boolean {
    if (model) {
      xmlStream.openNode(this.name);
      if (model.argb) {
        xmlStream.addAttribute('rgb', model.argb);
      } else if (model.theme !== undefined) {
        xmlStream.addAttribute('theme', model.theme);
        if (model.tint !== undefined) {
          xmlStream.addAttribute('tint', model.tint);
        }
      } else if (model.indexed !== undefined) {
        xmlStream.addAttribute('indexed', model.indexed);
      } else {
        xmlStream.addAttribute('auto', '1');
      }
      xmlStream.closeNode();
      return true;
    }
    return false;
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
