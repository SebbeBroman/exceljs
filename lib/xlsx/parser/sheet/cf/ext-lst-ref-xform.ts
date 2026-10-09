import type {ExtLstRefModel} from '../../../xform/sheet/cf/ext-lst-ref-xform.js';
export type {ExtLstRefModel} from '../../../xform/sheet/cf/ext-lst-ref-xform.js';
import BaseXform from '../../../base-parser.js';
import CompositeXform from '../../composite-xform.js';

class X14IdXform extends BaseXform<string> {
  override tag = 'x14:id';

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

  override createNewModel(): ExtLstRefModel {
    return {};
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    Object.assign(this.model as ExtLstRefModel, parser.model);
  }
}

export default ExtLstRefXform;
export {ExtLstRefXform};
