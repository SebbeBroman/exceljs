import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';
import VmlShapeXform from './vml-shape-xform.js';
import type {VmlShapeModel, VmlShapeRenderModel} from './vml-shape-xform.js';

export interface VmlNotesModel {
  comments: VmlShapeModel[];
  anchors?: Array<{br?: unknown; [key: string]: unknown}>;
}

// This class is (currently) single purposed to insert the triangle
// drawing icons on commented cells
class VmlNotesXform extends BaseXform<VmlNotesModel> {
  static DRAWING_ATTRIBUTES: Record<string, string> = {
    'xmlns:v': 'urn:schemas-microsoft-com:vml',
    'xmlns:o': 'urn:schemas-microsoft-com:office:office',
    'xmlns:x': 'urn:schemas-microsoft-com:office:excel',
  };

  override tag = 'xml';
  declare map: Record<string, BaseXform>;

  constructor() {
    super();
    this.map = {
      'v:shape': new VmlShapeXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: VmlNotesModel | null): void {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, VmlNotesXform.DRAWING_ATTRIBUTES);

    xmlStream.openNode('o:shapelayout', {'v:ext': 'edit'});
    xmlStream.leafNode('o:idmap', {'v:ext': 'edit', data: 1});
    xmlStream.closeNode();

    xmlStream.openNode('v:shapetype', {
      id: '_x0000_t202',
      coordsize: '21600,21600',
      'o:spt': 202,
      path: 'm,l,21600r21600,l21600,xe',
    });
    xmlStream.leafNode('v:stroke', {joinstyle: 'miter'});
    xmlStream.leafNode('v:path', {gradientshapeok: 't', 'o:connecttype': 'rect'});
    xmlStream.closeNode();

    model!.comments.forEach((item, index) => {
      this.map['v:shape'].render(
        xmlStream,
        item as unknown as VmlShapeRenderModel,
        index,
      );
    });

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
          comments: [],
        };
        break;
      default:
        this.parser = this.map[node.name];
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
        this.model!.comments.push(this.parser.model as VmlShapeModel);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  override reconcile(model?: VmlNotesModel | null, options?: XformOptions): void {
    model!.anchors!.forEach(anchor => {
      if (anchor.br) {
        this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
      } else {
        this.map['xdr:oneCellAnchor'].reconcile(anchor, options);
      }
    });
  }
}

export default VmlNotesXform;
export {VmlNotesXform};
