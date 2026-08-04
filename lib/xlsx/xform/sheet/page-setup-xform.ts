import _ from '../../../utils/under-dash.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface PageSetupModel {
  paperSize?: number;
  orientation?: string;
  horizontalDpi?: number;
  verticalDpi?: number;
  pageOrder?: string;
  blackAndWhite?: boolean;
  draft?: boolean;
  cellComments?: string;
  errors?: string;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  firstPageNumber?: number;
  useFirstPageNumber?: boolean;
  usePrinterDefaults?: boolean;
  copies?: number;
}

function booleanToXml(model: unknown): string | undefined {
  return model ? '1' : undefined;
}
function pageOrderToXml(model: string | undefined): string | undefined {
  switch (model) {
    case 'overThenDown':
      return model;
    default:
      return undefined;
  }
}
function cellCommentsToXml(model: string | undefined): string | undefined {
  switch (model) {
    case 'atEnd':
    case 'asDisplyed':
      return model;
    default:
      return undefined;
  }
}
function errorsToXml(model: string | undefined): string | undefined {
  switch (model) {
    case 'dash':
    case 'blank':
    case 'NA':
      return model;
    default:
      return undefined;
  }
}
function pageSizeToModel(value: string | undefined): number | undefined {
  return value !== undefined ? parseInt(value, 10) : undefined;
}

class PageSetupXform extends BaseXform<PageSetupModel> {
  override tag = 'pageSetup';

  override render(xmlStream: XmlStreamLike, model?: PageSetupModel | null): void {
    if (model) {
      const attributes = {
        paperSize: model.paperSize,
        orientation: model.orientation,
        horizontalDpi: model.horizontalDpi,
        verticalDpi: model.verticalDpi,
        pageOrder: pageOrderToXml(model.pageOrder),
        blackAndWhite: booleanToXml(model.blackAndWhite),
        draft: booleanToXml(model.draft),
        cellComments: cellCommentsToXml(model.cellComments),
        errors: errorsToXml(model.errors),
        scale: model.scale,
        fitToWidth: model.fitToWidth,
        fitToHeight: model.fitToHeight,
        firstPageNumber: model.firstPageNumber,
        useFirstPageNumber: booleanToXml(model.firstPageNumber),
        usePrinterDefaults: booleanToXml(model.usePrinterDefaults),
        copies: model.copies,
      };
      if (_.some(attributes, (value: unknown) => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }

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
