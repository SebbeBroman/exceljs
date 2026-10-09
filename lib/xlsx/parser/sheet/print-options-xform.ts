import type {PrintOptionsModel} from '../../xform/sheet/print-options-xform.js';
export type {PrintOptionsModel} from '../../xform/sheet/print-options-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class PrintOptionsXform extends BaseXform<PrintOptionsModel> {
  override tag = 'printOptions';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          showRowColHeaders: node.attributes.headings === '1',
          showGridLines: node.attributes.gridLines === '1',
          horizontalCentered: node.attributes.horizontalCentered === '1',
          verticalCentered: node.attributes.verticalCentered === '1',
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
}

export default PrintOptionsXform;
export {PrintOptionsXform};
