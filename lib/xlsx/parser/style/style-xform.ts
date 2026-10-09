import type {StyleModel} from '../../xform/style/style-xform.js';
export type {StyleModel} from '../../xform/style/style-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import AlignmentXform from './alignment-xform.js';
import type {AlignmentModel} from './alignment-xform.js';
import ProtectionXform from './protection-xform.js';
import type {ProtectionModel} from './protection-xform.js';

// Style assists translation from style model to/from xlsx
class StyleXform extends BaseXform<StyleModel> {
  xfId: boolean;
  declare map: {
    alignment: AlignmentXform;
    protection: ProtectionXform;
  };

  constructor(options?: {xfId?: boolean}) {
    super();

    this.xfId = !!(options && options.xfId);
    this.map = {
      alignment: new AlignmentXform(),
      protection: new ProtectionXform(),
    };
  }

  override tag = 'xf';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    // used during sax parsing of xml to build font object
    switch (node.name) {
      case 'xf':
        this.model = {
          numFmtId: parseInt(node.attributes.numFmtId, 10),
          fontId: parseInt(node.attributes.fontId, 10),
          fillId: parseInt(node.attributes.fillId, 10),
          borderId: parseInt(node.attributes.borderId, 10),
        };
        if (this.xfId) {
          this.model.xfId = parseInt(node.attributes.xfId, 10);
        }
        return true;
      case 'alignment':
        this.parser = this.map.alignment;
        this.parser.parseOpen(node);
        return true;
      case 'protection':
        this.parser = this.map.protection;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        if (this.map.protection === this.parser) {
          this.model!.protection = this.parser.model as ProtectionModel | null;
        } else {
          this.model!.alignment = this.parser.model as AlignmentModel | null;
        }
        this.parser = undefined;
      }
      return true;
    }
    return name !== 'xf';
  }
}

export default StyleXform;
export {StyleXform};
