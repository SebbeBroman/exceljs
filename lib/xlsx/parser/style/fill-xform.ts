import type {
  FillModel,
  GradientFillModel,
  PatternFillModel,
  StopModel,
} from '../../xform/style/fill-xform.js';
export type {
  FillModel,
  GradientFillModel,
  GradientCenter,
  PatternFillModel,
  StopModel,
} from '../../xform/style/fill-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import ColorXform from './color-xform.js';
import type {ColorModel} from './color-xform.js';

class StopXform extends BaseXform<StopModel> {
  declare map: {color: ColorXform};

  constructor() {
    super();

    this.map = {
      color: new ColorXform(),
    };
  }

  override tag = 'stop';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'stop':
        this.model = {
          position: parseFloat(node.attributes.position),
        };
        return true;
      case 'color':
        this.parser = this.map.color;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.color = this.parser.model as ColorModel;
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }
}

class PatternFillXform extends BaseXform<PatternFillModel> {
  declare map: {fgColor: ColorXform; bgColor: ColorXform};

  constructor() {
    super();

    this.map = {
      fgColor: new ColorXform('fgColor'),
      bgColor: new ColorXform('bgColor'),
    };
  }

  get name(): string {
    return 'pattern';
  }

  override tag = 'patternFill';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'patternFill':
        this.model = {
          type: 'pattern',
          pattern: node.attributes.patternType,
        };
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
        if (this.parser.model) {
          (this.model as Record<string, unknown>)[name as string] = this.parser.model;
        }
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }
}

class GradientFillXform extends BaseXform<GradientFillModel> {
  declare map: {stop: StopXform};

  constructor() {
    super();

    this.map = {
      stop: new StopXform(),
    };
  }

  get name(): string {
    return 'gradient';
  }

  override tag = 'gradientFill';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'gradientFill': {
        const model = (this.model = {
          stops: [] as StopModel[],
        } as GradientFillModel);
        if (node.attributes.degree) {
          model.gradient = 'angle';
          model.degree = parseInt(node.attributes.degree, 10);
        } else if (node.attributes.type === 'path') {
          model.gradient = 'path';
          model.center = {
            left: node.attributes.left ? parseFloat(node.attributes.left) : 0,
            top: node.attributes.top ? parseFloat(node.attributes.top) : 0,
          };
          if (node.attributes.right !== node.attributes.left) {
            model.center.right = node.attributes.right ? parseFloat(node.attributes.right) : 0;
          }
          if (node.attributes.bottom !== node.attributes.top) {
            model.center.bottom = node.attributes.bottom ? parseFloat(node.attributes.bottom) : 0;
          }
        }
        return true;
      }

      case 'stop':
        this.parser = this.map.stop;
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
        this.model!.stops.push(this.parser.model as StopModel);
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }
}

// Fill encapsulates translation from fill model to/from xlsx
class FillXform extends BaseXform<FillModel> {
  declare map: {
    patternFill: PatternFillXform;
    gradientFill: GradientFillXform;
  };

  constructor() {
    super();

    this.map = {
      patternFill: new PatternFillXform(),
      gradientFill: new GradientFillXform(),
    };
  }

  override tag = 'fill';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'fill':
        this.model = {} as FillModel;
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
        this.model = this.parser.model as FillModel;
        (this.model as {type?: string}).type = (
          this.parser as PatternFillXform | GradientFillXform
        ).name;
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }

  validStyle(value: string): boolean | undefined {
    return FillXform.validPatternValues[value];
  }

  static validPatternValues: Record<string, boolean> = [
    'none',
    'solid',
    'darkVertical',
    'darkGray',
    'mediumGray',
    'lightGray',
    'gray125',
    'gray0625',
    'darkHorizontal',
    'darkVertical',
    'darkDown',
    'darkUp',
    'darkGrid',
    'darkTrellis',
    'lightHorizontal',
    'lightVertical',
    'lightDown',
    'lightUp',
    'lightGrid',
    'lightTrellis',
    'lightGrid',
  ].reduce(
    (p, v) => {
      p[v] = true;
      return p;
    },
    {} as Record<string, boolean>,
  );

  static StopXform = StopXform;
  static PatternFillXform = PatternFillXform;
  static GradientFillXform = GradientFillXform;
}

export default FillXform;
export {FillXform};
