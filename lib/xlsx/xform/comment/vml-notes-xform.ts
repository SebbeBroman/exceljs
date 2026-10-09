import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
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
      this.map['v:shape'].render(xmlStream, item as unknown as VmlShapeRenderModel, index);
    });

    xmlStream.closeNode();
  }
}

export default VmlNotesXform;
export {VmlNotesXform};
