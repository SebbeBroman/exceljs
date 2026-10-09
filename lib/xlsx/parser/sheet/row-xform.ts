import type {
  RowXformOptions,
  RowXformConstructorOptions,
  RowXformModel,
} from '../../xform/sheet/row-xform.js';
export type {
  RowXformOptions,
  StylesRowLike,
  RowXformConstructorOptions,
  RowXformModel,
} from '../../xform/sheet/row-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import {parseBoolean} from '../../../utils/utils.js';
import CellXform from './cell-xform.js';
import type {CellXformModel, CellXformOptions} from './cell-xform.js';

class RowXform extends BaseXform<RowXformModel> {
  maxItems: number | undefined;
  override map: {c: CellXform};
  numRowsSeen = 0;

  constructor(options?: RowXformConstructorOptions) {
    super();

    this.maxItems = options && options.maxItems;
    this.map = {
      c: new CellXform(),
    };
  }

  override tag = 'row';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === 'row') {
      this.numRowsSeen += 1;
      const spans = node.attributes.spans
        ? node.attributes.spans.split(':').map(span => parseInt(span, 10))
        : [undefined, undefined];
      const model: RowXformModel = (this.model = {
        number: parseInt(node.attributes.r, 10),
        min: spans[0],
        max: spans[1],
        cells: [],
      });
      if (node.attributes.s) {
        model.styleId = parseInt(node.attributes.s, 10);
      }
      if (parseBoolean(node.attributes.hidden)) {
        model.hidden = true;
      }
      if (parseBoolean(node.attributes.bestFit)) {
        model.bestFit = true;
      }
      if (node.attributes.ht) {
        model.height = parseFloat(node.attributes.ht);
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (parseBoolean(node.attributes.collapsed)) {
        model.collapsed = true;
      }
      return true;
    }

    this.parser = this.map[node.name as 'c'];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.cells.push(this.parser.model as CellXformModel);
        if (this.maxItems && this.model!.cells.length > this.maxItems) {
          throw new Error(`Max column count (${this.maxItems}) exceeded`);
        }
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }

  override reconcile(model?: RowXformModel | null, options?: RowXformOptions): void {
    if (!model || !options) {
      return;
    }
    model.style = model.styleId ? options.styles.getStyleModel(model.styleId) : {};
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      cellXform.reconcile(cellModel, options as CellXformOptions);
    });
  }
}

export default RowXform;
export {RowXform};
