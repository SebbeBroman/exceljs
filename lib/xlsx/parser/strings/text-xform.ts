import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

//   <t xml:space="preserve"> is </t>

class TextXform extends BaseXform<string> {
  _text: string[] = [];

  override tag = 't';

  override get model(): string {
    return this._text
      .join('')
      .replace(/_x([0-9A-F]{4})_/g, ($0, $1: string) => String.fromCharCode(parseInt($1, 16)));
  }

  override set model(value: string | null | undefined) {
    // BaseXform.reset() assigns null; clear buffer so subsequent parses start clean
    if (value == null) {
      this._text = [];
    }
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 't':
        this._text = [];
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    this._text.push(text);
  }

  override parseClose(): boolean {
    return false;
  }
}

export default TextXform;
export {TextXform};
