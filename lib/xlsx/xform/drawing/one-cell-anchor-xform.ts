import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import StaticXform from '../static-xform.js';
import BaseCellAnchorXform from './base-cell-anchor-xform.js';
import type {BaseCellAnchorModel, PictureRef} from './base-cell-anchor-xform.js';
import CellPositionXform from './cell-position-xform.js';
import type {CellPositionModel} from './cell-position-xform.js';
import ExtXform from './ext-xform.js';
import type {ExtModel} from './ext-xform.js';
import PicXform from './pic-xform.js';
import type {PicModel} from './pic-xform.js';

export interface OneCellAnchorRange {
  editAs?: string;
  tl?: CellPositionModel;
  ext?: ExtModel;
}

export interface OneCellAnchorModel extends BaseCellAnchorModel {
  range: OneCellAnchorRange;
  picture?: PicModel & PictureRef;
}

class OneCellAnchorXform extends BaseCellAnchorXform<OneCellAnchorModel> {
  override tag = 'xdr:oneCellAnchor';
  declare map: {
    'xdr:from': CellPositionXform;
    'xdr:ext': ExtXform;
    'xdr:pic': PicXform;
    'xdr:clientData': StaticXform;
  };

  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:ext': new ExtXform({tag: 'xdr:ext'}),
      'xdr:pic': new PicXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  override prepare(model?: OneCellAnchorModel | null, options?: XformOptions): void {
    this.map['xdr:pic'].prepare(model!.picture, options);
  }

  override render(xmlStream: XmlStreamLike, model?: OneCellAnchorModel | null): void {
    xmlStream.openNode(this.tag, {editAs: model!.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model!.range.tl);
    this.map['xdr:ext'].render(xmlStream, model!.range.ext);
    this.map['xdr:pic'].render(xmlStream, model!.picture);
    this.map['xdr:clientData'].render(xmlStream);

    xmlStream.closeNode();
  }
}

export default OneCellAnchorXform;
export {OneCellAnchorXform};
