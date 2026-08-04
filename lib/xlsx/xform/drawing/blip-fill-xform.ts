import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
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

  override render(xmlStream: XmlStreamLike, model?: BlipModel | null): void {
    xmlStream.openNode(this.tag);

    this.map['a:blip'].render(xmlStream, model);

    // TODO: options for this + parsing
    xmlStream.openNode('a:stretch');
    xmlStream.leafNode('a:fillRect');
    xmlStream.closeNode();

    xmlStream.closeNode();
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
