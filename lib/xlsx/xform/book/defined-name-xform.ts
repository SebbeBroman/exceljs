import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import colCache from '../../../utils/col-cache.js';

export interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
}

class DefinedNamesXform extends BaseXform<DefinedNameModel> {
  _parsedName?: string;
  _parsedLocalSheetId?: string;
  _parsedText?: string[];

  override render(xmlStream: XmlStreamLike, model?: DefinedNameModel | null): void {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName', {
      name: model!.name,
      localSheetId: model!.localSheetId,
    });
    xmlStream.writeText(model!.ranges.join(','));
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'definedName':
        this._parsedName = node.attributes.name;
        this._parsedLocalSheetId = node.attributes.localSheetId;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    this._parsedText!.push(text);
  }

  override parseClose(): boolean {
    this.model = {
      name: this._parsedName!,
      ranges: extractRanges(this._parsedText!.join('')),
    };
    if (this._parsedLocalSheetId !== undefined) {
      this.model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    return false;
  }
}

function isValidRange(range: string): boolean {
  try {
    colCache.decodeEx(range);
    return true;
  } catch (_err) {
    return false;
  }
}

function extractRanges(parsedText: string): string[] {
  const ranges: string[] = [];
  let quotesOpened = false;
  let last = '';
  parsedText.split(',').forEach(item => {
    if (!item) {
      return;
    }
    const quotes = (item.match(/'/g) || []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += `${item},`;
      } else if (isValidRange(item)) {
        ranges.push(item);
      }
      return;
    }
    const quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven && isValidRange(item)) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (isValidRange(last + item)) {
        ranges.push(last + item);
      }
      last = '';
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  });
  return ranges;
}

export default DefinedNamesXform;
export {DefinedNamesXform};
