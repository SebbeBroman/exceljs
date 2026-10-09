import type {DxfModel} from '../../xform/style/dxf-xform.js';
export type {DxfModel} from '../../xform/style/dxf-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import AlignmentXform from './alignment-xform.js';
import BorderXform from './border-xform.js';
import FillXform from './fill-xform.js';
import FontXform from './font-xform.js';
import NumFmtXform from './numfmt-xform.js';
import ProtectionXform from './protection-xform.js';

// Style assists translation from style model to/from xlsx
class DxfXform extends BaseXform<DxfModel> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: any;

  constructor() {
    super();

    this.map = {
      alignment: new AlignmentXform(),
      border: new BorderXform(),
      fill: new FillXform(),
      font: new FontXform(),
      numFmt: new NumFmtXform(),
      protection: new ProtectionXform(),
    };
  }

  override tag = 'dxf';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        // this node is often repeated. Need to reset children
        this.reset();
        return true;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (name === this.tag) {
      this.model = {
        alignment: this.map.alignment.model,
        border: this.map.border.model,
        fill: this.map.fill.model,
        font: this.map.font.model,
        numFmt: this.map.numFmt.model as unknown as string,
        protection: this.map.protection.model,
      };
      return false;
    }

    return true;
  }
}

export default DxfXform;
export {DxfXform};
