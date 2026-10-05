import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import VmlTextboxXform from './vml-textbox-xform.js';
import type {VmlTextboxRenderModel} from './vml-textbox-xform.js';
import VmlClientDataXform from './vml-client-data-xform.js';
import type {VmlClientDataRenderModel} from './vml-client-data-xform.js';

export interface VmlShapeModel {
  margins: {
    insetmode?: string;
    inset?: number[];
  };
  anchor?: string | number[];
  editAs?: string;
  protection?: {
    locked?: string;
    lockText?: string;
  };
}

export type VmlShapeRenderModel = VmlTextboxRenderModel & VmlClientDataRenderModel;

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

  static V_SHAPE_ATTRIBUTES = (
    model: VmlShapeRenderModel,
    index: number,
  ): Record<string, unknown> => ({
    id: `_x0000_s${1025 + index}`,
    type: '#_x0000_t202',
    style:
      'position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden',
    fillcolor: 'infoBackground [80]',
    strokecolor: 'none [81]',
    'o:insetmode': model.note.margins && model.note.margins.insetmode,
  });

  override render(xmlStream: XmlStreamLike, model?: VmlShapeModel | null, index?: number): void {
    const renderModel = model as unknown as VmlShapeRenderModel;
    xmlStream.openNode('v:shape', VmlShapeXform.V_SHAPE_ATTRIBUTES(renderModel, index!));

    xmlStream.leafNode('v:fill', {color2: 'infoBackground [80]'});
    xmlStream.leafNode('v:shadow', {color: 'none [81]', obscured: 't'});
    xmlStream.leafNode('v:path', {'o:connecttype': 'none'});
    this.map['v:textbox'].render(xmlStream, renderModel as never);
    this.map['x:ClientData'].render(xmlStream, renderModel as never);

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
