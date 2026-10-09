import type {
  BaseCellAnchorModel,
  PictureRef,
  DrawingReconcileOptions,
} from '../../xform/drawing/base-cell-anchor-xform.js';
export type {
  BaseCellAnchorModel,
  AnchorRange,
  PictureRef,
  DrawingReconcileOptions,
  DrawingRel,
} from '../../xform/drawing/base-cell-anchor-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class BaseCellAnchorXform<
  TModel extends BaseCellAnchorModel = BaseCellAnchorModel,
> extends BaseXform<TModel> {
  declare map: Record<string, BaseXform>;

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          range: {
            editAs: node.attributes.editAs || 'oneCell',
          },
        } as TModel;
        break;
      default:
        this.parser = this.map[node.name];
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

  reconcilePicture(
    model: PictureRef | null | undefined,
    options: DrawingReconcileOptions,
  ): unknown {
    if (model && model.rId) {
      const rel = options.rels[model.rId];
      const match = rel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
      if (match) {
        const name = match[1];
        const mediaId = options.mediaIndex[name];
        return options.media[mediaId];
      }
    }
    return undefined;
  }
}

export default BaseCellAnchorXform;
export {BaseCellAnchorXform};
