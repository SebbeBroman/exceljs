import BaseXform from '../../base-xform.js';
import type {XmlStreamLike} from '../../base-xform.js';

class FormulaXform extends BaseXform<string> {
  override tag = 'formula';

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

export default FormulaXform;
export {FormulaXform};
