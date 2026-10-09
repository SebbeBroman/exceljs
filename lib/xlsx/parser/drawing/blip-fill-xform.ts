import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import BlipXform from './blip-xform.js';
import type {BlipModel} from './blip-xform.js';

class BlipFillXform extends BaseXform<BlipModel> {
  override tag = 'xdr:blipFill';
  declare map: {
    'a:blip': BlipXform;
  };

  constructor() {
    super();

    this.map = {
      'a:blip': new BlipXform(),
    };
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        break;

      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model = this.map['a:blip'].model as BlipModel;
        return false;

      default:
        return true;
    }
  }
}

export default BlipFillXform;
export {BlipFillXform};
