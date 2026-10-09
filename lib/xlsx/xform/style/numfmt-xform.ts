import defaultNumFormats from '../../defaultnumformats.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface NumFmtModel {
  id: number;
  formatCode: string;
}

function hashDefaultFormats(): Record<string, number> {
  const hash: Record<string, number> = {};
  for (const [id, dnf] of Object.entries(defaultNumFormats as Record<string, {f?: string}>)) {
    if (dnf.f) {
      hash[dnf.f] = parseInt(id, 10);
    }
    // at some point, add the other cultures here...
  }
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
