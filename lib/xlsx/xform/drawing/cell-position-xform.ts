import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import IntegerXform from '../simple/integer-xform.js';

export interface CellPositionModel {
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;
}

export interface CellPositionXformOptions {
  tag: string;
}

class CellPositionXform extends BaseXform<CellPositionModel> {
  override tag: string;
  declare map: {
    'xdr:col': IntegerXform;
    'xdr:colOff': IntegerXform;
    'xdr:row': IntegerXform;
    'xdr:rowOff': IntegerXform;
  };

  constructor(options: CellPositionXformOptions) {
    super();

    this.tag = options.tag;
    this.map = {
      'xdr:col': new IntegerXform({tag: 'xdr:col', zero: true}),
      'xdr:colOff': new IntegerXform({tag: 'xdr:colOff', zero: true}),
      'xdr:row': new IntegerXform({tag: 'xdr:row', zero: true}),
      'xdr:rowOff': new IntegerXform({tag: 'xdr:rowOff', zero: true}),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CellPositionModel | null): void {
    xmlStream.openNode(this.tag);

    this.map['xdr:col'].render(xmlStream, model!.nativeCol);
    this.map['xdr:colOff'].render(xmlStream, model!.nativeColOff);

    this.map['xdr:row'].render(xmlStream, model!.nativeRow);
    this.map['xdr:rowOff'].render(xmlStream, model!.nativeRowOff);

    xmlStream.closeNode();
  }
}

export default CellPositionXform;
export {CellPositionXform};
