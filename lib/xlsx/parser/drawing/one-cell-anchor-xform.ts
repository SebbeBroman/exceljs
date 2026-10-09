import type {OneCellAnchorModel} from '../../xform/drawing/one-cell-anchor-xform.js';
export type {
  OneCellAnchorModel,
  OneCellAnchorRange,
} from '../../xform/drawing/one-cell-anchor-xform.js';
import type {XformOptions} from '../../base-parser.js';
import StaticXform from '../static-xform.js';
import BaseCellAnchorXform from './base-cell-anchor-xform.js';
import type {DrawingReconcileOptions, PictureRef} from './base-cell-anchor-xform.js';
import CellPositionXform from './cell-position-xform.js';
import type {CellPositionModel} from './cell-position-xform.js';
import ExtXform from './ext-xform.js';
import type {ExtModel} from './ext-xform.js';
import PicXform from './pic-xform.js';
import type {PicModel} from './pic-xform.js';

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
        this.model!.range.ext = this.map['xdr:ext'].model as ExtModel;
        this.model!.picture = this.map['xdr:pic'].model as PicModel & PictureRef;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  override reconcile(model?: OneCellAnchorModel | null, options?: XformOptions): void {
    model!.medium = this.reconcilePicture(model!.picture, options as DrawingReconcileOptions);
  }
}

export default OneCellAnchorXform;
export {OneCellAnchorXform};
