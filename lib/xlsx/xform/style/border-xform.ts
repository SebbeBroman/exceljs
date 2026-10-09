import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import ColorXform from './color-xform.js';
import type {ColorModel} from './color-xform.js';

export interface EdgeModel {
  style?: string;
  color?: ColorModel;
  up?: boolean;
  down?: boolean;
}

export interface BorderModel {
  color?: ColorModel;
  top?: EdgeModel;
  left?: EdgeModel;
  bottom?: EdgeModel;
  right?: EdgeModel;
  diagonal?: EdgeModel;
}

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

  override render(
    xmlStream: XmlStreamLike,
    model?: EdgeModel | null,
    defaultColor?: ColorModel,
  ): void {
    const color = (model && model.color) || defaultColor || this.defaultColor;
    xmlStream.openNode(this.name);
    if (model && model.style) {
      xmlStream.addAttribute('style', model.style);
      if (color) {
        this.map.color.render(xmlStream, color);
      }
    }
    xmlStream.closeNode();
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

  override render(xmlStream: XmlStreamLike, model?: BorderModel | null): void {
    const {color} = model!;
    xmlStream.openNode('border');
    if (model!.diagonal && model!.diagonal.style) {
      if (model!.diagonal.up) {
        xmlStream.addAttribute('diagonalUp', '1');
      }
      if (model!.diagonal.down) {
        xmlStream.addAttribute('diagonalDown', '1');
      }
    }
    const add = (edgeModel: EdgeModel | undefined, edgeXform: EdgeXform): void => {
      let edge = edgeModel;
      if (edge && !edge.color && model!.color) {
        // don't mess with incoming models
        edge = {
          ...edge,
          color: model!.color,
        };
      }
      edgeXform.render(xmlStream, edge, color);
    };
    add(model!.left, this.map.left);
    add(model!.right, this.map.right);
    add(model!.top, this.map.top);
    add(model!.bottom, this.map.bottom);
    add(model!.diagonal, this.map.diagonal);

    xmlStream.closeNode();
  }
}

export default BorderXform;
export {BorderXform, EdgeXform};
