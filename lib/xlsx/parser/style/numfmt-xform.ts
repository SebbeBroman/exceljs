import type {NumFmtModel} from '../../xform/style/numfmt-xform.js';
export type {NumFmtModel} from '../../xform/style/numfmt-xform.js';
import defaultNumFormats from '../../defaultnumformats.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

// NumFmt encapsulates translation between number format and xlsx
class NumFmtXform extends BaseXform<NumFmtModel> {
  id: number | undefined;
  formatCode: string | undefined;

  constructor(id?: number, formatCode?: string) {
    super();

    this.id = id;
    this.formatCode = formatCode;
  }

  override tag = 'numFmt';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'numFmt':
        this.model = {
          id: parseInt(node.attributes.numFmtId, 10),
          formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1'),
        };
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }

  static getDefaultFmtCode(numFmtId: number | string): string | undefined {
    const entry = (defaultNumFormats as Record<string, {f?: string}>)[numFmtId];
    return entry && entry.f;
  }
}

export default NumFmtXform;
export {NumFmtXform};
