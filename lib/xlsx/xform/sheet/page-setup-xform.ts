import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
      if (Object.values(attributes).some(value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }
}

export default PageSetupXform;
export {PageSetupXform};
