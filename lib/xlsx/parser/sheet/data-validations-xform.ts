import type {
  DataValidationsModel,
  DataValidationEntry,
} from '../../xform/sheet/data-validations-xform.js';
export type {
  DataValidationsModel,
  DataValidationEntry,
} from '../../xform/sheet/data-validations-xform.js';
import {excelToDate, parseBoolean} from '../../../utils/utils.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import Range from '../../../model/range.js';
import {subtractRectangle} from '../../../model/rectangles.js';

function assign(
  definedName: DataValidationEntry,
  attributes: Record<string, string>,
  name: string,
  defaultValue?: string,
): void {
  const value = attributes[name];
  if (value !== undefined) {
    definedName[name] = value;
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}

function assignBool(
  definedName: DataValidationEntry,
  attributes: Record<string, string>,
  name: string,
  defaultValue?: boolean,
): void {
  const value = attributes[name];
  if (value !== undefined) {
    definedName[name] = parseBoolean(value);
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}

class DataValidationsXform extends BaseXform<DataValidationsModel> {
  _address?: string;
  _dataValidation?: DataValidationEntry;
  _formula?: string[];
  private rangeKeys = new Set<string>();

  override tag = 'dataValidations';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'dataValidations':
        this.model = {};
        this.rangeKeys.clear();
        return true;

      case 'dataValidation': {
        this._address = node.attributes.sqref;
        const dataValidation: DataValidationEntry = {
          type: node.attributes.type || 'any',
          formulae: [],
        };

        if (node.attributes.type) {
          assignBool(dataValidation, node.attributes, 'allowBlank');
        }
        assignBool(dataValidation, node.attributes, 'showInputMessage');
        assignBool(dataValidation, node.attributes, 'showErrorMessage');

        switch (dataValidation.type) {
          case 'any':
          case 'list':
          case 'custom':
            break;
          default:
            assign(dataValidation, node.attributes, 'operator', 'between');
            break;
        }
        assign(dataValidation, node.attributes, 'promptTitle');
        assign(dataValidation, node.attributes, 'prompt');
        assign(dataValidation, node.attributes, 'errorStyle');
        assign(dataValidation, node.attributes, 'errorTitle');
        assign(dataValidation, node.attributes, 'error');

        this._dataValidation = dataValidation;
        return true;
      }

      case 'formula1':
      case 'formula2':
        this._formula = [];
        return true;

      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this._formula) {
      this._formula.push(text);
    }
  }

  override parseClose(name?: string): boolean {
    switch (name) {
      case 'dataValidations':
        return false;
      case 'dataValidation': {
        if (!this._dataValidation!.formulae || !this._dataValidation!.formulae.length) {
          delete this._dataValidation!.formulae;
          delete this._dataValidation!.operator;
        }
        // The four known cases: 1. E4:L9 N4:U9  2.E4 L9  3. N4:U9  4. E4
        const list = this._address!.split(/\s+/g) || [];
        list.filter(Boolean).forEach(addr => {
          const rect = new Range(addr).model;
          // Remove covered portions of prior rules without enumerating their cells.
          const addresses =
            rect.top === rect.bottom && rect.left === rect.right
              ? [...this.rangeKeys]
              : Object.keys(this.model!);
          for (const address of addresses) {
            const rule = this.model![address];
            const previous = new Range(address).model;
            const pieces = subtractRectangle(previous, rect);
            if (pieces.length === 1 && pieces[0] === previous) continue;
            delete this.model![address];
            this.rangeKeys.delete(address);
            for (const piece of pieces) {
              const key = new Range(piece).shortRange;
              this.model![key] = rule;
              if (key.includes(':')) this.rangeKeys.add(key);
            }
          }
          const key = new Range(rect).shortRange;
          this.model![key] = this._dataValidation!;
          if (key.includes(':')) this.rangeKeys.add(key);
        });
        return true;
      }
      case 'formula1':
      case 'formula2': {
        let formula: unknown = this._formula!.join('');
        switch (this._dataValidation!.type) {
          case 'whole':
          case 'textLength':
            formula = parseInt(formula as string, 10);
            break;
          case 'decimal':
            formula = parseFloat(formula as string);
            break;
          case 'date':
            formula = excelToDate(parseFloat(formula as string));
            break;
          default:
            break;
        }
        this._dataValidation!.formulae!.push(formula);
        this._formula = undefined;
        return true;
      }
      default:
        return true;
    }
  }
}

export default DataValidationsXform;
export {DataValidationsXform};
