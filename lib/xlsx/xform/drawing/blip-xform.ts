import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default BlipXform;
export {BlipXform};
