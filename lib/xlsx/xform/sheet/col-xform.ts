import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';

export interface ColStyleModel {
  [key: string]: unknown;
}

export interface ColModel {
  min: number;
  max: number;
  width?: number;
  styleId?: number;
  style?: ColStyleModel;
  hidden?: boolean;
  bestFit?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface StylesLike {
  addStyleModel(style: ColStyleModel): number | undefined;
  getStyleModel(styleId: number): ColStyleModel | undefined;
}

export interface ColXformOptions extends XformOptions {
  styles: StylesLike;
}

class ColXform extends BaseXform<ColModel> {
  override tag = 'col';

  override prepare(model?: ColModel | null, options?: ColXformOptions): void {
    if (!model || !options) {
      return;
    }
    const styleId = options.styles.addStyleModel(model.style || {});
    if (styleId) {
      model.styleId = styleId;
    }
  }

  override render(xmlStream: XmlStreamLike, model?: ColModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode('col');
    xmlStream.addAttribute('min', model.min);
    xmlStream.addAttribute('max', model.max);
    if (model.width) {
      xmlStream.addAttribute('width', model.width);
    }
    if (model.styleId) {
      xmlStream.addAttribute('style', model.styleId);
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden', '1');
    }
    if (model.bestFit) {
      xmlStream.addAttribute('bestFit', '1');
    }
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel', model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed', '1');
    }
    xmlStream.addAttribute('customWidth', '1');
    xmlStream.closeNode();
  }
}

export default ColXform;
export {ColXform};
