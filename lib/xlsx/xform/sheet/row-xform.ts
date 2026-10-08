import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from '../base-xform.js';
import {parseBoolean} from '../../../utils/utils.js';
import CellXform from './cell-xform.js';
import type {CellXformModel, CellXformOptions} from './cell-xform.js';

export interface RowXformModel {
  number: number;
  min?: number;
  max?: number;
  cells: CellXformModel[];
  height?: number;
  hidden?: boolean;
  bestFit?: boolean;
  styleId?: number;
  style?: unknown;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface RowXformConstructorOptions {
  maxItems?: number;
}

export interface StylesRowLike {
  addStyleModel(style: unknown): number | undefined;
  getStyleModel(styleId: number): unknown;
}

export interface RowXformOptions extends XformOptions {
  styles: StylesRowLike;
}

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

  override prepare(model?: RowXformModel | null, options?: RowXformOptions): void {
    if (!model || !options) {
      return;
    }
    const styleId = options.styles.addStyleModel(model.style);
    if (styleId) {
      model.styleId = styleId;
    }
    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      cellXform.prepare(cellModel, options as CellXformOptions);
    });
  }

  override render(xmlStream: XmlStreamLike, model?: RowXformModel | null, options?: unknown): void {
    if (!model) {
      return;
    }
    xmlStream.openNode('row');
    xmlStream.addAttribute('r', model.number);
    if (model.height) {
      xmlStream.addAttribute('ht', model.height);
      xmlStream.addAttribute('customHeight', '1');
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden', '1');
    }
    if (model.min! > 0 && model.max! > 0 && model.min! <= model.max!) {
      xmlStream.addAttribute('spans', `${model.min}:${model.max}`);
    }
    if (model.styleId) {
      xmlStream.addAttribute('s', model.styleId);
      xmlStream.addAttribute('customFormat', '1');
    }
    xmlStream.addAttribute('x14ac:dyDescent', '0.25');
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel', model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed', '1');
    }

    const cellXform = this.map.c;
    model.cells.forEach(cellModel => {
      // original passed options as 3rd arg; CellXform.render only uses (stream, model)
      void options;
      cellXform.render(xmlStream, cellModel);
    });

    xmlStream.closeNode();
  }

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
