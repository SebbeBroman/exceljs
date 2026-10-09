import colCache from '../../../utils/col-cache.js';
import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import TwoCellAnchorXform from './two-cell-anchor-xform.js';
import type {TwoCellAnchorModel} from './two-cell-anchor-xform.js';
import OneCellAnchorXform from './one-cell-anchor-xform.js';
import type {OneCellAnchorModel} from './one-cell-anchor-xform.js';

export type DrawingAnchorModel = (OneCellAnchorModel | TwoCellAnchorModel) & {
  anchorType?: 'xdr:twoCellAnchor' | 'xdr:oneCellAnchor';
  range: OneCellAnchorModel['range'] | TwoCellAnchorModel['range'] | string;
};

export interface DrawingModel {
  anchors: DrawingAnchorModel[];
}

function getAnchorType(model: DrawingAnchorModel): 'xdr:twoCellAnchor' | 'xdr:oneCellAnchor' {
  const range =
    typeof model.range === 'string'
      ? (colCache.decode(model.range) as {br?: unknown})
      : model.range;

  return range && (range as {br?: unknown}).br ? 'xdr:twoCellAnchor' : 'xdr:oneCellAnchor';
}

class DrawingXform extends BaseXform<DrawingModel> {
  static DRAWING_ATTRIBUTES: Record<string, string> = {
    'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  };

  override tag = 'xdr:wsDr';
  declare map: {
    'xdr:twoCellAnchor': TwoCellAnchorXform;
    'xdr:oneCellAnchor': OneCellAnchorXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:twoCellAnchor': new TwoCellAnchorXform(),
      'xdr:oneCellAnchor': new OneCellAnchorXform(),
    };
  }

  override prepare(model?: DrawingModel | null): void {
    model!.anchors.forEach((item, index) => {
      item.anchorType = getAnchorType(item);
      const anchor = this.map[item.anchorType];
      anchor.prepare(item as never, {index});
    });
  }

  override render(xmlStream: XmlStreamLike, model?: DrawingModel | null): void {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, DrawingXform.DRAWING_ATTRIBUTES);

    model!.anchors.forEach(item => {
      const anchor = this.map[item.anchorType!];
      anchor.render(xmlStream, item as never);
    });

    xmlStream.closeNode();
  }
}

export default DrawingXform;
export {DrawingXform};
