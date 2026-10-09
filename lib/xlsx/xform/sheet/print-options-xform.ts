import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface PrintOptionsModel {
  showRowColHeaders?: boolean;
  showGridLines?: boolean;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
}

function booleanToXml(model: boolean | undefined): string | undefined {
  return model ? '1' : undefined;
}

class PrintOptionsXform extends BaseXform<PrintOptionsModel> {
  override tag = 'printOptions';

  override render(xmlStream: XmlStreamLike, model?: PrintOptionsModel | null): void {
    if (model) {
      const attributes = {
        headings: booleanToXml(model.showRowColHeaders),
        gridLines: booleanToXml(model.showGridLines),
        horizontalCentered: booleanToXml(model.horizontalCentered),
        verticalCentered: booleanToXml(model.verticalCentered),
      };
      if (Object.values(attributes).some(value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }
}

export default PrintOptionsXform;
export {PrintOptionsXform};
