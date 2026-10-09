import type {VmlNotesModel} from '../../xform/comment/vml-notes-xform.js';
export type {VmlNotesModel} from '../../xform/comment/vml-notes-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode, XformOptions} from '../../base-parser.js';
import VmlShapeXform from './vml-shape-xform.js';
import type {VmlShapeModel} from './vml-shape-xform.js';

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
