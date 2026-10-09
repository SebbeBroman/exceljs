import type {BorderModel, EdgeModel} from '../../xform/style/border-xform.js';
export type {BorderModel, EdgeModel} from '../../xform/style/border-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import {parseBoolean} from '../../../utils/utils.js';
import ColorXform from './color-xform.js';
import type {ColorModel} from './color-xform.js';

class EdgeXform extends BaseXform<EdgeModel> {
  name: string;
  declare map: {color: ColorXform};
  defaultColor?: ColorModel;

  constructor(name: string) {
    super();

    this.name = name;
    this.tag = name;
    this.map = {
      color: new ColorXform(),
    };
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.name: {
        const {style} = node.attributes;
        if (style) {
          this.model = {
            style,
          };
        } else {
          this.model = undefined;
        }
        return true;
      }
      case 'color':
        this.parser = this.map.color;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }

    if (name === this.name) {
      if (this.map.color.model) {
        if (!this.model) {
          this.model = {};
        }
        this.model.color = this.map.color.model;
      }
    }

    return false;
  }

  validStyle(value: string): boolean | undefined {
    return EdgeXform.validStyleValues[value];
  }

  static validStyleValues: Record<string, boolean> = [
    'thin',
    'dashed',
    'dotted',
    'dashDot',
    'hair',
    'dashDotDot',
    'slantDashDot',
    'mediumDashed',
    'mediumDashDotDot',
    'mediumDashDot',
    'medium',
    'double',
    'thick',
  ].reduce(
    (p, v) => {
      p[v] = true;
      return p;
    },
    {} as Record<string, boolean>,
  );
}

// Border encapsulates translation from border model to/from xlsx
class BorderXform extends BaseXform<BorderModel> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: any;
  diagonalUp?: boolean;
  diagonalDown?: boolean;

  constructor() {
    super();

    this.map = {
      top: new EdgeXform('top'),
      left: new EdgeXform('left'),
      bottom: new EdgeXform('bottom'),
      right: new EdgeXform('right'),
      diagonal: new EdgeXform('diagonal'),
    };
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'border':
        this.reset();
        this.diagonalUp = parseBoolean(node.attributes.diagonalUp);
        this.diagonalDown = parseBoolean(node.attributes.diagonalDown);
        return true;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (name === 'border') {
      const model = (this.model = {} as BorderModel);
      const add = function (
        key: keyof BorderModel,
        edgeModel: EdgeModel | null | undefined,
        extensions?: {up?: boolean; down?: boolean},
      ): void {
        if (edgeModel) {
          if (extensions) {
            Object.assign(edgeModel, extensions);
          }
          (model as Record<string, unknown>)[key] = edgeModel;
        }
      };
      add('left', this.map.left.model);
      add('right', this.map.right.model);
      add('top', this.map.top.model);
      add('bottom', this.map.bottom.model);
      add('diagonal', this.map.diagonal.model, {up: this.diagonalUp, down: this.diagonalDown});
    }
    return false;
  }
}

export default BorderXform;
export {BorderXform, EdgeXform};
