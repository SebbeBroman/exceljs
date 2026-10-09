import type {SheetFormatPropertiesModel} from '../../xform/sheet/sheet-format-properties-xform.js';
export type {SheetFormatPropertiesModel} from '../../xform/sheet/sheet-format-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class SheetFormatPropertiesXform extends BaseXform<SheetFormatPropertiesModel> {
  override tag = 'sheetFormatPr';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'sheetFormatPr') {
      this.model = {
        defaultRowHeight: parseFloat(node.attributes.defaultRowHeight || '0'),
        dyDescent: parseFloat(node.attributes['x14ac:dyDescent'] || '0'),
        outlineLevelRow: parseInt(node.attributes.outlineLevelRow || '0', 10),
        outlineLevelCol: parseInt(node.attributes.outlineLevelCol || '0', 10),
      };
      if (node.attributes.defaultColWidth) {
        this.model.defaultColWidth = parseFloat(node.attributes.defaultColWidth);
      }
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default SheetFormatPropertiesXform;
export {SheetFormatPropertiesXform};
