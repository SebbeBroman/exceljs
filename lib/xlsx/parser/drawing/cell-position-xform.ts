import type {
  CellPositionXformOptions,
  CellPositionModel,
} from '../../xform/drawing/cell-position-xform.js';
export type {
  CellPositionXformOptions,
  CellPositionModel,
} from '../../xform/drawing/cell-position-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import IntegerXform from '../simple/integer-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
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
      case this.tag:
        this.model = {
          nativeCol: this.map['xdr:col'].model as number,
          nativeColOff: this.map['xdr:colOff'].model as number,
          nativeRow: this.map['xdr:row'].model as number,
          nativeRowOff: this.map['xdr:rowOff'].model as number,
        };
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

export default CellPositionXform;
export {CellPositionXform};
