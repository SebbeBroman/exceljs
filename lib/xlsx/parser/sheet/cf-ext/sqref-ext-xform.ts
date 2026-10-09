import BaseXform from '../../../base-parser.js';

class SqrefExtXform extends BaseXform<string> {
  override tag = 'xm:sqref';

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

export default SqrefExtXform;
export {SqrefExtXform};
