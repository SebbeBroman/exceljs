import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import CNvPrXform from './c-nv-pr-xform.js';
import type {CNvPrModel} from './c-nv-pr-xform.js';
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

  override render(xmlStream: XmlStreamLike, model?: CNvPrModel | null): void {
    xmlStream.openNode(this.tag);
    this.map['xdr:cNvPr'].render(xmlStream, model);
    this.map['xdr:cNvPicPr'].render(xmlStream);
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
        this.model = this.map['xdr:cNvPr'].model as HLinkClickModel;
        return false;
      default:
        return true;
    }
  }
}

export default NvPicPrXform;
export {NvPicPrXform};
