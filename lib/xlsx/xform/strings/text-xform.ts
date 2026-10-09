import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

//   <t xml:space="preserve"> is </t>

class TextXform extends BaseXform<string> {
  override tag = 't';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.openNode('t');
    if (/^\s|\n|\s$/.test(model as string)) {
      xmlStream.addAttribute('xml:space', 'preserve');
    }
    xmlStream.writeText(model);
    xmlStream.closeNode();
  }
}

export default TextXform;
export {TextXform};
