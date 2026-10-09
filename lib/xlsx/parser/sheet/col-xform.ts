import type {ColXformOptions, ColModel} from '../../xform/sheet/col-xform.js';
export type {
  ColXformOptions,
  StylesLike,
  ColModel,
  ColStyleModel,
} from '../../xform/sheet/col-xform.js';
import {parseBoolean} from '../../../utils/utils.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class ColXform extends BaseXform<ColModel> {
  override tag = 'col';

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'col') {
      const model: ColModel = (this.model = {
        min: parseInt(node.attributes.min || '0', 10),
        max: parseInt(node.attributes.max || '0', 10),
        width:
          node.attributes.width === undefined
            ? undefined
            : parseFloat(node.attributes.width || '0'),
      });
      if (node.attributes.style) {
        model.styleId = parseInt(node.attributes.style, 10);
      }
      if (parseBoolean(node.attributes.hidden)) {
        model.hidden = true;
      }
      if (parseBoolean(node.attributes.bestFit)) {
        model.bestFit = true;
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (parseBoolean(node.attributes.collapsed)) {
        model.collapsed = true;
      }
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }

  override reconcile(model?: ColModel | null, options?: ColXformOptions): void {
    // reconcile column styles
    if (model?.styleId && options) {
      model.style = options.styles.getStyleModel(model.styleId);
    }
  }
}

export default ColXform;
export {ColXform};
