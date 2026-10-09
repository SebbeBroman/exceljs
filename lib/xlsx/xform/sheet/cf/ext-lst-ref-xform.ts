import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

export interface ExtLstRefModel {
  x14Id?: string;
}

class X14IdXform extends BaseXform<string> {
  override tag = 'x14:id';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }
}

class ExtXform extends CompositeXform<ExtLstRefModel> {
  idXform: X14IdXform;

  constructor() {
    super();

    this.map = {
      'x14:id': (this.idXform = new X14IdXform()),
    };
  }

  override tag = 'ext';

  override render(xmlStream: XmlStreamLike, model?: ExtLstRefModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode(this.tag, {
      uri: '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });

    this.idXform.render(xmlStream, model.x14Id);

    xmlStream.closeNode();
  }
}

class ExtLstRefXform extends CompositeXform<ExtLstRefModel> {
  constructor() {
    super();
    this.map = {
      ext: new ExtXform(),
    };
  }

  override tag = 'extLst';

  override render(xmlStream: XmlStreamLike, model?: ExtLstRefModel | null): void {
    xmlStream.openNode(this.tag);
    this.map.ext.render(xmlStream, model);
    xmlStream.closeNode();
  }
}

export default ExtLstRefXform;
export {ExtLstRefXform};
