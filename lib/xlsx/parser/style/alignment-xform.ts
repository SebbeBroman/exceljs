import type {AlignmentModel} from '../../xform/style/alignment-xform.js';
export type {AlignmentModel} from '../../xform/style/alignment-xform.js';
import Enums from '../../../model/enums.js';
import {parseBoolean, validInt} from '../../../utils/utils.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

const validation = {
  horizontalValues: [
    'left',
    'center',
    'right',
    'fill',
    'centerContinuous',
    'distributed',
    'justify',
  ].reduce(
    (p, v) => {
      p[v] = true;
      return p;
    },
    {} as Record<string, boolean>,
  ),
  horizontal(value: string | undefined): string | undefined {
    return this.horizontalValues[value as string] ? value : undefined;
  },

  verticalValues: ['top', 'middle', 'bottom', 'distributed', 'justify'].reduce(
    (p, v) => {
      p[v] = true;
      return p;
    },
    {} as Record<string, boolean>,
  ),
  vertical(value: string | undefined): string | undefined {
    if (value === 'middle') return 'center';
    return this.verticalValues[value as string] ? value : undefined;
  },
  wrapText(value: boolean | undefined): true | undefined {
    return value ? true : undefined;
  },
  shrinkToFit(value: boolean | undefined): true | undefined {
    return value ? true : undefined;
  },
  textRotation(value: number | string | undefined): number | string | undefined {
    switch (value) {
      case 'vertical':
        return value;
      default: {
        const intVal = validInt(value);
        return intVal >= -90 && intVal <= 90 ? intVal : undefined;
      }
    }
  },
  indent(value: number | undefined): number {
    const intVal = validInt(value);
    return Math.max(0, intVal);
  },
  readingOrder(value: string | undefined): number | undefined {
    switch (value) {
      case 'ltr':
        return (Enums as {ReadingOrder: {LeftToRight: number; RightToLeft: number}}).ReadingOrder
          .LeftToRight;
      case 'rtl':
        return (Enums as {ReadingOrder: {LeftToRight: number; RightToLeft: number}}).ReadingOrder
          .RightToLeft;
      default:
        return undefined;
    }
  },
};

const textRotationXform = {
  toXml(textRotation: number | string | undefined): number | undefined {
    textRotation = validation.textRotation(textRotation);
    if (textRotation) {
      if (textRotation === 'vertical') {
        return 255;
      }

      const tr = Math.round(textRotation as number);
      if (tr >= 0 && tr <= 90) {
        return tr;
      }

      if (tr < 0 && tr >= -90) {
        return 90 - tr;
      }
    }
    return undefined;
  },
  toModel(textRotation: string | undefined): number | string | undefined {
    const tr = validInt(textRotation);
    if (tr !== undefined) {
      if (tr === 255) {
        return 'vertical';
      }
      if (tr >= 0 && tr <= 90) {
        return tr;
      }
      if (tr > 90 && tr <= 180) {
        return 90 - tr;
      }
    }
    return undefined;
  },
};

// Alignment encapsulates translation from style.alignment model to/from xlsx
class AlignmentXform extends BaseXform<AlignmentModel | null> {
  override tag = 'alignment';

  override parseOpen(node: XmlNode): void {
    const model: AlignmentModel = {};

    let valid = false;
    function add(truthy: unknown, name: keyof AlignmentModel, value: unknown): void {
      if (truthy) {
        (model as Record<string, unknown>)[name] = value;
        valid = true;
      }
    }
    add(node.attributes.horizontal, 'horizontal', node.attributes.horizontal);
    add(
      node.attributes.vertical,
      'vertical',
      node.attributes.vertical === 'center' ? 'middle' : node.attributes.vertical,
    );
    add(node.attributes.wrapText, 'wrapText', parseBoolean(node.attributes.wrapText));
    add(node.attributes.shrinkToFit, 'shrinkToFit', parseBoolean(node.attributes.shrinkToFit));
    add(node.attributes.indent, 'indent', parseInt(node.attributes.indent, 10));
    add(
      node.attributes.textRotation,
      'textRotation',
      textRotationXform.toModel(node.attributes.textRotation),
    );
    add(
      node.attributes.readingOrder,
      'readingOrder',
      node.attributes.readingOrder === '2' ? 'rtl' : 'ltr',
    );

    this.model = valid ? model : null;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default AlignmentXform;
export {AlignmentXform};
