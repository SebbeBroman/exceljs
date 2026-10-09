import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import StaticXform from '../static-xform.js';
import BaseCellAnchorXform from './base-cell-anchor-xform.js';
import type {BaseCellAnchorModel, PictureRef} from './base-cell-anchor-xform.js';
import CellPositionXform from './cell-position-xform.js';
import type {CellPositionModel} from './cell-position-xform.js';
import PicXform from './pic-xform.js';
import type {PicModel} from './pic-xform.js';

export interface TwoCellAnchorRange {
  editAs?: string;
  tl?: CellPositionModel;
  br?: CellPositionModel;
}

export interface TwoCellAnchorModel extends BaseCellAnchorModel {
  range: TwoCellAnchorRange;
  picture?: PicModel & PictureRef;
}

class TwoCellAnchorXform extends BaseCellAnchorXform<TwoCellAnchorModel> {
  override tag = 'xdr:twoCellAnchor';
  declare map: {
    'xdr:from': CellPositionXform;
    'xdr:to': CellPositionXform;
    'xdr:pic': PicXform;
    'xdr:clientData': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:to': new CellPositionXform({tag: 'xdr:to'}),
      'xdr:pic': new PicXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  override prepare(model?: TwoCellAnchorModel | null, options?: XformOptions): void {
    this.map['xdr:pic'].prepare(model!.picture, options);
  }

  override render(xmlStream: XmlStreamLike, model?: TwoCellAnchorModel | null): void {
    xmlStream.openNode(this.tag, {editAs: model!.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model!.range.tl);
    this.map['xdr:to'].render(xmlStream, model!.range.br);
    this.map['xdr:pic'].render(xmlStream, model!.picture);
    this.map['xdr:clientData'].render(xmlStream);

    xmlStream.closeNode();
  }
}

export default TwoCellAnchorXform;
export {TwoCellAnchorXform};
