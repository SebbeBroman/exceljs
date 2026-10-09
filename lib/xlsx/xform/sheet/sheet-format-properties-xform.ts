import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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

      if (Object.values(attributes).some(value => value !== undefined)) {
        xmlStream.leafNode('sheetFormatPr', attributes);
      }
    }
  }
}

export default SheetFormatPropertiesXform;
export {SheetFormatPropertiesXform};
