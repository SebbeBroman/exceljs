import BaseXform from '../../../base-parser.js';

class FExtXform extends BaseXform<string> {
  override tag = 'xm:f';

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

export default FExtXform;
export {FExtXform};
