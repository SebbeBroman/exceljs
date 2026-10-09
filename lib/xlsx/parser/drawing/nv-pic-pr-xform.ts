import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import CNvPrXform from './c-nv-pr-xform.js';
import CNvPicPrXform from './c-nv-pic-pr-xform.js';
import type {HLinkClickModel} from './hlink-click-xform.js';

class NvPicPrXform extends BaseXform<HLinkClickModel> {
  override tag = 'xdr:nvPicPr';
  declare map: {
    'xdr:cNvPr': CNvPrXform;
    'xdr:cNvPicPr': CNvPicPrXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:cNvPr': new CNvPrXform(),
      'xdr:cNvPicPr': new CNvPicPrXform(),
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
        this.model = this.map['xdr:cNvPr'].model as HLinkClickModel;
        return false;
      default:
        return true;
    }
  }
}

export default NvPicPrXform;
export {NvPicPrXform};
