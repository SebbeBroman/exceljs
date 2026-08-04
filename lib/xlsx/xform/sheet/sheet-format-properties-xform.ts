import _ from '../../../utils/under-dash.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface SheetFormatPropertiesModel {
  defaultRowHeight?: number;
  defaultColWidth?: number;
  outlineLevelRow?: number;
  outlineLevelCol?: number;
  dyDescent?: number;
}

class SheetFormatPropertiesXform extends BaseXform<SheetFormatPropertiesModel> {
  override tag = 'sheetFormatPr';

  override render(xmlStream: XmlStreamLike, model?: SheetFormatPropertiesModel | null): void {
    if (model) {
      const attributes: Record<string, unknown> = {
        defaultRowHeight: model.defaultRowHeight,
        outlineLevelRow: model.outlineLevelRow,
        outlineLevelCol: model.outlineLevelCol,
        'x14ac:dyDescent': model.dyDescent,
      };
      if (model.defaultColWidth) {
        attributes.defaultColWidth = model.defaultColWidth;
      }

      // default value for 'defaultRowHeight' is 15, this should not be 'custom'
      if (!model.defaultRowHeight || model.defaultRowHeight !== 15) {
        attributes.customHeight = '1';
      }

      if (_.some(attributes, (value: unknown) => value !== undefined)) {
        xmlStream.leafNode('sheetFormatPr', attributes);
      }
    }
  }

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
