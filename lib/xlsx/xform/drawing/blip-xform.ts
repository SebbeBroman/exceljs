import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface BlipModel {
  rId: string;
}

class BlipXform extends BaseXform<BlipModel> {
  override tag = 'a:blip';

  override render(xmlStream: XmlStreamLike, model?: BlipModel | null): void {
    xmlStream.leafNode(this.tag, {
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:embed': model!.rId,
      cstate: 'print',
    });
    // TODO: handle children (e.g. a:extLst=>a:ext=>a14:useLocalDpi
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:embed'],
        };
        return true;
      default:
        return true;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
}

export default BlipXform;
export {BlipXform};
