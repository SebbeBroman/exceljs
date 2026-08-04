import colCache from '../utils/col-cache.js';
import Anchor from './anchor.js';
import type {ImageHyperlinkValue} from '../../index.js';
import type {AnchorAddress, AnchorWorksheet, AnchorModel} from './anchor.js';
import type {RangeAddress} from '../utils/col-cache.js';

export type ImageType = 'background' | 'image';

export interface ImageRangeExt {
  width?: number;
  height?: number;
}

export interface ImageRangeModel {
  tl: AnchorAddress | string;
  br?: AnchorAddress | string | null;
  ext?: ImageRangeExt;
  editAs?: string;
  hyperlinks?: ImageHyperlinkValue;
}

export interface ImageInternalRange {
  tl: Anchor;
  br?: Anchor | null | false;
  ext?: ImageRangeExt;
  editAs?: string;
  hyperlinks?: ImageHyperlinkValue;
}

export interface ImageModelInput {
  type: ImageType;
  imageId: number;
  range?: string | ImageRangeModel;
  hyperlinks?: ImageHyperlinkValue;
}

export interface ImageBackgroundModel {
  type: 'background';
  imageId: number;
}

export interface ImageRangeOutputModel {
  type: 'image';
  imageId: number;
  hyperlinks?: ImageHyperlinkValue;
  range: {
    tl: AnchorModel;
    br?: AnchorModel;
    ext?: ImageRangeExt;
    editAs?: string;
  };
}

class Image {
  worksheet: AnchorWorksheet;
  type!: ImageType;
  imageId!: number;
  range!: ImageInternalRange;

  constructor(worksheet: AnchorWorksheet, model: ImageModelInput) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model(): ImageBackgroundModel | ImageRangeOutputModel {
    switch (this.type) {
      case 'background':
        return {
          type: this.type,
          imageId: this.imageId,
        };
      case 'image':
        return {
          type: this.type,
          imageId: this.imageId,
          hyperlinks: this.range.hyperlinks,
          range: {
            tl: this.range.tl.model,
            // Preserve original short-circuit (`br && br.model` may be false)
            br: (this.range.br && this.range.br.model) as AnchorModel | undefined,
            ext: this.range.ext,
            editAs: this.range.editAs,
          },
        };
      default:
        throw new Error('Invalid Image Type');
    }
  }

  set model({type, imageId, range, hyperlinks}: ImageModelInput) {
    this.type = type;
    this.imageId = imageId;

    if (type === 'image') {
      if (typeof range === 'string') {
        const decoded = colCache.decode(range) as RangeAddress;
        this.range = {
          tl: new Anchor(this.worksheet, {col: decoded.left, row: decoded.top}, -1),
          br: new Anchor(this.worksheet, {col: decoded.right, row: decoded.bottom}, 0),
          editAs: 'oneCell',
        };
      } else if (range) {
        this.range = {
          tl: new Anchor(this.worksheet, range.tl as AnchorAddress, 0),
          br: (range.br && new Anchor(this.worksheet, range.br as AnchorAddress, 0)) || undefined,
          ext: range.ext,
          editAs: range.editAs,
          hyperlinks: hyperlinks || range.hyperlinks,
        };
      }
    }
  }
}

export default Image;
export {Image};
