import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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
