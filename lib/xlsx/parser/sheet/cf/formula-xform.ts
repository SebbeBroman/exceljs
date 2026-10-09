import BaseXform from '../../../base-parser.js';

class FormulaXform extends BaseXform<string> {
  override tag = 'formula';

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
