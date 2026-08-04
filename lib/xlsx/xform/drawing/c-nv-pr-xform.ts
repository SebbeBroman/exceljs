import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import HlickClickXform from './hlink-click-xform.js';
import type {HLinkClickModel} from './hlink-click-xform.js';
import ExtLstXform from './ext-lst-xform.js';

export interface CNvPrModel extends HLinkClickModel {
  index?: number;
}

class CNvPrXform extends BaseXform<HLinkClickModel> {
  override tag = 'xdr:cNvPr';
  declare map: {
    'a:hlinkClick': HlickClickXform;
    'a:extLst': ExtLstXform;
  };

  constructor() {
    super();

    this.map = {
      'a:hlinkClick': new HlickClickXform(),
      'a:extLst': new ExtLstXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CNvPrModel | null): void {
    xmlStream.openNode(this.tag, {
      id: model!.index,
      name: `Picture ${model!.index}`,
    });
    this.map['a:hlinkClick'].render(xmlStream, model);
    this.map['a:extLst'].render(xmlStream);
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
        this.model = this.map['a:hlinkClick'].model as HLinkClickModel;
        return false;
      default:
        return true;
    }
  }
}

export default CNvPrXform;
export {CNvPrXform};
