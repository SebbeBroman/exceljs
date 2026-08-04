import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

//   <t xml:space="preserve"> is </t>

class TextXform extends BaseXform<string> {
  _text: string[] = [];

  override tag = 't';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.openNode('t');
    if (/^\s|\n|\s$/.test(model as string)) {
      xmlStream.addAttribute('xml:space', 'preserve');
    }
    xmlStream.writeText(model);
    xmlStream.closeNode();
  }

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
