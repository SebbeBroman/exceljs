import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
const EMU_PER_PIXEL_AT_96_DPI = 9525;

export interface ExtModel {
  width: number;
  height: number;
}

export interface ExtXformOptions {
  tag: string;
}

class ExtXform extends BaseXform<ExtModel> {
  override tag: string;

  constructor(options: ExtXformOptions) {
    super();

    this.tag = options.tag;
    this.map = {};
  }

  override render(xmlStream: XmlStreamLike, model?: ExtModel | null): void {
    xmlStream.openNode(this.tag);

    const width = Math.floor(model!.width * EMU_PER_PIXEL_AT_96_DPI);
    const height = Math.floor(model!.height * EMU_PER_PIXEL_AT_96_DPI);

    xmlStream.addAttribute('cx', width);
    xmlStream.addAttribute('cy', height);

    xmlStream.closeNode();
  }
}

export default ExtXform;
export {ExtXform};
