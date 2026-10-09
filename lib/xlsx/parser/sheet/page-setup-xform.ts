import type {PageSetupModel} from '../../xform/sheet/page-setup-xform.js';
export type {PageSetupModel} from '../../xform/sheet/page-setup-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
function pageSizeToModel(value: string | undefined): number | undefined {
  return value !== undefined ? parseInt(value, 10) : undefined;
}

class PageSetupXform extends BaseXform<PageSetupModel> {
  override tag = 'pageSetup';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          paperSize: pageSizeToModel(node.attributes.paperSize),
          orientation: node.attributes.orientation || 'portrait',
          horizontalDpi: parseInt(node.attributes.horizontalDpi || '4294967295', 10),
          verticalDpi: parseInt(node.attributes.verticalDpi || '4294967295', 10),
          pageOrder: node.attributes.pageOrder || 'downThenOver',
          blackAndWhite: node.attributes.blackAndWhite === '1',
          draft: node.attributes.draft === '1',
          cellComments: node.attributes.cellComments || 'None',
          errors: node.attributes.errors || 'displayed',
          scale: parseInt(node.attributes.scale || '100', 10),
          fitToWidth: parseInt(node.attributes.fitToWidth || '1', 10),
          fitToHeight: parseInt(node.attributes.fitToHeight || '1', 10),
          firstPageNumber: parseInt(node.attributes.firstPageNumber || '1', 10),
          useFirstPageNumber: node.attributes.useFirstPageNumber === '1',
          usePrinterDefaults: node.attributes.usePrinterDefaults === '1',
          copies: parseInt(node.attributes.copies || '1', 10),
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

export default PageSetupXform;
export {PageSetupXform};
