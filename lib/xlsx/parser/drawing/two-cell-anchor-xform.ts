import type {TwoCellAnchorModel} from '../../xform/drawing/two-cell-anchor-xform.js';
export type {
  TwoCellAnchorModel,
  TwoCellAnchorRange,
} from '../../xform/drawing/two-cell-anchor-xform.js';
import type {XformOptions} from '../../base-parser.js';
import StaticXform from '../static-xform.js';
import BaseCellAnchorXform from './base-cell-anchor-xform.js';
import type {DrawingReconcileOptions, PictureRef} from './base-cell-anchor-xform.js';
import CellPositionXform from './cell-position-xform.js';
import type {CellPositionModel} from './cell-position-xform.js';
import PicXform from './pic-xform.js';
import type {PicModel} from './pic-xform.js';

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

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model!.range.tl = this.map['xdr:from'].model as CellPositionModel;
        this.model!.range.br = this.map['xdr:to'].model as CellPositionModel;
        this.model!.picture = this.map['xdr:pic'].model as PicModel & PictureRef;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  override reconcile(model?: TwoCellAnchorModel | null, options?: XformOptions): void {
    model!.medium = this.reconcilePicture(model!.picture, options as DrawingReconcileOptions);
  }
}

export default TwoCellAnchorXform;
export {TwoCellAnchorXform};
