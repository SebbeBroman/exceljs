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

  override parseOpen(): void {
    this.model = '';
  }

  override parseText(text: string): void {
    this.model = (this.model || '') + text;
  }

  override parseClose(name?: string): boolean {
    return name !== this.tag;
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

  override createNewModel(): ExtLstRefModel {
    return {};
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    (this.model as ExtLstRefModel).x14Id = parser.model as string;
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

  override createNewModel(): ExtLstRefModel {
    return {};
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    Object.assign(this.model as ExtLstRefModel, parser.model);
  }
}

export default ExtLstRefXform;
export {ExtLstRefXform};
