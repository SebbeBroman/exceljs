import type {ExtXformOptions, ExtModel} from '../../xform/drawing/ext-xform.js';
export type {ExtXformOptions, ExtModel} from '../../xform/drawing/ext-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
const EMU_PER_PIXEL_AT_96_DPI = 9525;

class ExtXform extends BaseXform<ExtModel> {
  override tag: string;

  constructor(options: ExtXformOptions) {
    super();

    this.tag = options.tag;
    this.map = {};
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === this.tag) {
      this.model = {
        width: parseInt(node.attributes.cx || '0', 10) / EMU_PER_PIXEL_AT_96_DPI,
        height: parseInt(node.attributes.cy || '0', 10) / EMU_PER_PIXEL_AT_96_DPI,
      };
      return true;
    }
    return false;
  }

  override parseText(_text?: string): void {}

  override parseClose(_name?: string): boolean {
    return false;
  }
}

export default ExtXform;
export {ExtXform};
