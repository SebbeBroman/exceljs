import _ from '../../../utils/under-dash.js';
import defaultNumFormats from '../../defaultnumformats.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface NumFmtModel {
  id: number;
  formatCode: string;
}

function hashDefaultFormats(): Record<string, number> {
  const hash: Record<string, number> = {};
  _.each(defaultNumFormats, (dnf: {f?: string}, id: string | number) => {
    if (dnf.f) {
      hash[dnf.f] = parseInt(String(id), 10);
    }
    // at some point, add the other cultures here...
  });
  return hash;
}
const defaultFmtHash = hashDefaultFormats();

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

  override render(xmlStream: XmlStreamLike, model?: NumFmtModel | null): void {
    xmlStream.leafNode('numFmt', {numFmtId: model!.id, formatCode: model!.formatCode});
  }

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

  static getDefaultFmtId(formatCode: string): number | undefined {
    return defaultFmtHash[formatCode];
  }

  static getDefaultFmtCode(numFmtId: number | string): string | undefined {
    const entry = (defaultNumFormats as Record<string, {f?: string}>)[numFmtId];
    return entry && entry.f;
  }
}

export default NumFmtXform;
export {NumFmtXform};
