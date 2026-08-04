import colCache from '../../../utils/col-cache.js';
import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';
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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchors: [],
        };
        break;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.anchors.push(this.parser.model as DrawingAnchorModel);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  override reconcile(model?: DrawingModel | null, options?: XformOptions): void {
    model!.anchors.forEach(anchor => {
      if (anchor.br) {
        this.map['xdr:twoCellAnchor'].reconcile(anchor as TwoCellAnchorModel, options);
      } else {
        this.map['xdr:oneCellAnchor'].reconcile(anchor as OneCellAnchorModel, options);
      }
    });
  }
}

export default DrawingXform;
export {DrawingXform};
