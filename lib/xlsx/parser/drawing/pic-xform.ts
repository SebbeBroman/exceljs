import type {PicModel} from '../../xform/drawing/pic-xform.js';
export type {PicModel} from '../../xform/drawing/pic-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import StaticXform from '../static-xform.js';
import BlipFillXform from './blip-fill-xform.js';
import NvPicPrXform from './nv-pic-pr-xform.js';

class PicXform extends BaseXform<PicModel> {
  override tag = 'xdr:pic';
  declare map: {
    'xdr:nvPicPr': NvPicPrXform;
    'xdr:blipFill': BlipFillXform;
    'xdr:spPr': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:nvPicPr': new NvPicPrXform(),
      'xdr:blipFill': new BlipFillXform(),
      'xdr:spPr': new StaticXform({tag: 'xdr:spPr'}),
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
        this.mergeModel((this.parser.model || {}) as Record<string, unknown>);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

export default PicXform;
export {PicXform};
