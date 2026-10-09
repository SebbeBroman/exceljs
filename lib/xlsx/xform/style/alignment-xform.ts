import Enums from '../../../model/enums.js';
import {validInt} from '../../../utils/utils.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface AlignmentModel {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number | string;
  readingOrder?: string;
}

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

  override render(xmlStream: XmlStreamLike, model?: AlignmentModel | null): void {
    xmlStream.addRollback();
    xmlStream.openNode('alignment');

    let isValid = false;
    function add(name: string, value: unknown): void {
      if (value) {
        xmlStream.addAttribute(name, value);
        isValid = true;
      }
    }
    add('horizontal', validation.horizontal(model!.horizontal));
    add('vertical', validation.vertical(model!.vertical));
    add('wrapText', validation.wrapText(model!.wrapText) ? '1' : false);
    add('shrinkToFit', validation.shrinkToFit(model!.shrinkToFit) ? '1' : false);
    add('indent', validation.indent(model!.indent));
    add('textRotation', textRotationXform.toXml(model!.textRotation));
    add('readingOrder', validation.readingOrder(model!.readingOrder));

    xmlStream.closeNode();

    if (isValid) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  }
}

export default AlignmentXform;
export {AlignmentXform};
