import type {VmlShapeModel} from '../../xform/comment/vml-shape-xform.js';
export type {VmlShapeRenderModel, VmlShapeModel} from '../../xform/comment/vml-shape-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import VmlTextboxXform from './vml-textbox-xform.js';
import VmlClientDataXform from './vml-client-data-xform.js';

class VmlShapeXform extends BaseXform<VmlShapeModel> {
  override tag = 'v:shape';
  declare map: {
    'v:textbox': VmlTextboxXform;
    'x:ClientData': VmlClientDataXform;
  };

  constructor() {
    super();
    this.map = {
      'v:textbox': new VmlTextboxXform(),
      'x:ClientData': new VmlClientDataXform(),
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
        this.model = {
          margins: {
            insetmode: node.attributes['o:insetmode'],
          },
          anchor: '',
          editAs: '',
          protection: {},
        };
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

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag: {
        const textboxModel = this.map['v:textbox'].model;
        const clientDataModel = this.map['x:ClientData'].model;
        this.model!.margins.inset = textboxModel?.inset;
        this.model!.protection = clientDataModel?.protection;
        this.model!.anchor = clientDataModel?.anchor;
        this.model!.editAs = clientDataModel?.editAs;
        return false;
      }
      default:
        return true;
    }
  }
}

export default VmlShapeXform;
export {VmlShapeXform};
