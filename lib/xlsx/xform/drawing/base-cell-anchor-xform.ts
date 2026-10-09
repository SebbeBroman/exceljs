import BaseXform from '../base-xform.js';
import type {XformOptions} from '../base-xform.js';

export interface DrawingRel {
  Target: string;
  [key: string]: unknown;
}

export interface DrawingReconcileOptions extends XformOptions {
  rels: Record<string, DrawingRel>;
  mediaIndex: Record<string, number>;
  media: unknown[];
}

export interface PictureRef {
  rId?: string;
  [key: string]: unknown;
}

export interface AnchorRange {
  editAs?: string;
  tl?: unknown;
  br?: unknown;
  ext?: unknown;
}

export interface BaseCellAnchorModel {
  range: AnchorRange;
  picture?: PictureRef;
  medium?: unknown;
  br?: unknown;
  [key: string]: unknown;
}

class BaseCellAnchorXform<
  TModel extends BaseCellAnchorModel = BaseCellAnchorModel,
> extends BaseXform<TModel> {
  declare map: Record<string, BaseXform>;
}

export default BaseCellAnchorXform;
export {BaseCellAnchorXform};
