import TextXform from './text-xform.js';
import RichTextXform from './rich-text-xform.js';
import type {RichTextModel} from './rich-text-xform.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    const {name} = node;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (name === this.tag) {
      this.model = {
        sb: parseInt(node.attributes.sb, 10),
        eb: parseInt(node.attributes.eb, 10),
      };
      return true;
    }
    this.parser = this.map[name as keyof typeof this.map];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        switch (name) {
          case 'r': {
            let rt = (this.model as PhoneticTextModel).richText;
            if (!rt) {
              rt = (this.model as PhoneticTextModel).richText = [];
            }
            rt.push(this.parser.model as RichTextModel);
            break;
          }
          case 't':
            (this.model as PhoneticTextModel).text = this.parser.model as string;
            break;
          default:
            break;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export default PhoneticTextXform;
export {PhoneticTextXform};
