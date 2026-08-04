import TextXform from './text-xform.js';
import RichTextXform from './rich-text-xform.js';
import type {RichTextModel} from './rich-text-xform.js';
import PhoneticTextXform from './phonetic-text-xform.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

export interface SharedStringRichModel {
  richText?: RichTextModel[];
}

export type SharedStringModel = string | SharedStringRichModel | null | undefined;

class SharedStringXform extends BaseXform<SharedStringModel> {
  declare map: {
    r: RichTextXform;
    t: TextXform;
    rPh: PhoneticTextXform;
  };

  constructor(model?: SharedStringModel) {
    super();

    this.model = model;

    this.map = {
      r: new RichTextXform(),
      t: new TextXform(),
      rPh: new PhoneticTextXform(),
    };
  }

  override tag = 'si';

  override render(xmlStream: XmlStreamLike, model?: SharedStringModel | null): void {
    xmlStream.openNode(this.tag);
    if (model && typeof model === 'object' && Object.hasOwn(model, 'richText') && model.richText) {
      if (model.richText.length) {
        model.richText.forEach(text => {
          this.map.r.render(xmlStream, text);
        });
      } else {
        this.map.t.render(xmlStream, '');
      }
    } else if (model !== undefined && model !== null) {
      this.map.t.render(xmlStream, model as string);
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
      this.model = {};
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
            const m = this.model as SharedStringRichModel;
            let rt = m.richText;
            if (!rt) {
              rt = m.richText = [];
            }
            rt.push(this.parser.model as RichTextModel);
            break;
          }
          case 't':
            this.model = this.parser.model as string;
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

export default SharedStringXform;
export {SharedStringXform};
