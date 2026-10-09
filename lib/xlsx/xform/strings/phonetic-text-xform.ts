import TextXform from './text-xform.js';
import RichTextXform from './rich-text-xform.js';
import type {RichTextModel} from './rich-text-xform.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

// <rPh sb="0" eb="1">
//   <t>(its pronounciation in KATAKANA)</t>
// </rPh>

export interface PhoneticTextModel {
  sb?: number;
  eb?: number;
  text?: string;
  richText?: RichTextModel[];
}

class PhoneticTextXform extends BaseXform<PhoneticTextModel> {
  declare map: {
    r: RichTextXform;
    t: TextXform;
  };

  constructor() {
    super();

    this.map = {
      r: new RichTextXform(),
      t: new TextXform(),
    };
  }

  override tag = 'rPh';

  override render(xmlStream: XmlStreamLike, model?: PhoneticTextModel | null): void {
    xmlStream.openNode(this.tag, {
      sb: model?.sb || 0,
      eb: model?.eb || 0,
    });
    if (model && Object.hasOwn(model, 'richText') && model.richText) {
      const {r} = this.map;
      model.richText.forEach(text => {
        r.render(xmlStream, text);
      });
    } else if (model) {
      this.map.t.render(xmlStream, model.text);
    }
    xmlStream.closeNode();
  }
}

export default PhoneticTextXform;
export {PhoneticTextXform};
