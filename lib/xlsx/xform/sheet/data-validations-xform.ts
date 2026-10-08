import {isEqual} from '../../../utils/object.js';
import utils from '../../../utils/utils.js';
import colCache from '../../../utils/col-cache.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import Range from '../../../model/range.js';

export interface DataValidationEntry {
  type: string;
  formulae?: unknown[];
  operator?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  showErrorMessage?: boolean;
  errorStyle?: string;
  errorTitle?: string;
  error?: string;
  sqref?: string;
  [key: string]: unknown;
}

export type DataValidationsModel = Record<string, DataValidationEntry>;

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
    definedName[name] = utils.parseBoolean(value);
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}

function optimiseDataValidations(model: DataValidationsModel): DataValidationEntry[] {
  // Squeeze alike data validations together into rectangular ranges
  // to reduce file size and speed up Excel load time
  const dvList = Object.keys(model)
    .map(address => ({
      address,
      dataValidation: model[address],
      marked: false,
    }))
    .sort((a, b) => (a.address < b.address ? -1 : a.address > b.address ? 1 : 0));
  const dvMap = Object.fromEntries(dvList.map(dv => [dv.address, dv]));
  const matchCol = (
    addr: {row: number; col: number; address: string},
    height: number,
    col: number,
  ): boolean => {
    for (let i = 0; i < height; i++) {
      const otherAddress = colCache.encodeAddress(addr.row + i, col);
      if (!model[otherAddress] || !isEqual(model[addr.address], model[otherAddress])) {
        return false;
      }
    }
    return true;
  };
  return dvList
    .map(dv => {
      if (!dv.marked) {
        const addr = colCache.decodeEx(dv.address) as {
          dimensions?: string;
          row: number;
          col: number;
          address?: string;
        };
        if (addr.dimensions) {
          dvMap[addr.dimensions].marked = true;
          return {
            ...dv.dataValidation,
            sqref: dv.address,
          };
        }

        // iterate downwards - finding matching cells
        let height = 1;
        let otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        while (model[otherAddress] && isEqual(dv.dataValidation, model[otherAddress])) {
          height++;
          otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        }

        // iterate rightwards...

        let width = 1;
        while (
          matchCol({row: addr.row, col: addr.col, address: dv.address}, height, addr.col + width)
        ) {
          width++;
        }

        // mark all included addresses
        for (let i = 0; i < height; i++) {
          for (let j = 0; j < width; j++) {
            otherAddress = colCache.encodeAddress(addr.row + i, addr.col + j);
            dvMap[otherAddress].marked = true;
          }
        }

        if (height > 1 || width > 1) {
          const bottom = addr.row + (height - 1);
          const right = addr.col + (width - 1);
          return {
            ...dv.dataValidation,
            sqref: `${dv.address}:${colCache.encodeAddress(bottom, right)}`,
          };
        }
        return {
          ...dv.dataValidation,
          sqref: dv.address,
        };
      }
      return null;
    })
    .filter(Boolean) as DataValidationEntry[];
}

class DataValidationsXform extends BaseXform<DataValidationsModel> {
  _address?: string;
  _dataValidation?: DataValidationEntry;
  _formula?: string[];

  override tag = 'dataValidations';

  override render(xmlStream: XmlStreamLike, model?: DataValidationsModel | null): void {
    if (!model) {
      return;
    }
    const optimizedModel = optimiseDataValidations(model);
    if (optimizedModel.length) {
      xmlStream.openNode('dataValidations', {count: optimizedModel.length});

      optimizedModel.forEach(value => {
        xmlStream.openNode('dataValidation');

        if (value.type !== 'any') {
          xmlStream.addAttribute('type', value.type);

          if (value.operator && value.type !== 'list' && value.operator !== 'between') {
            xmlStream.addAttribute('operator', value.operator);
          }
          if (value.allowBlank) {
            xmlStream.addAttribute('allowBlank', '1');
          }
        }
        if (value.showInputMessage) {
          xmlStream.addAttribute('showInputMessage', '1');
        }
        if (value.promptTitle) {
          xmlStream.addAttribute('promptTitle', value.promptTitle);
        }
        if (value.prompt) {
          xmlStream.addAttribute('prompt', value.prompt);
        }
        if (value.showErrorMessage) {
          xmlStream.addAttribute('showErrorMessage', '1');
        }
        if (value.errorStyle) {
          xmlStream.addAttribute('errorStyle', value.errorStyle);
        }
        if (value.errorTitle) {
          xmlStream.addAttribute('errorTitle', value.errorTitle);
        }
        if (value.error) {
          xmlStream.addAttribute('error', value.error);
        }
        xmlStream.addAttribute('sqref', value.sqref);
        (value.formulae || []).forEach((formula, index) => {
          xmlStream.openNode(`formula${index + 1}`);
          if (value.type === 'date') {
            xmlStream.writeText(utils.dateToExcel(new Date(formula as string | number | Date)));
          } else {
            xmlStream.writeText(formula);
          }
          xmlStream.closeNode();
        });
        xmlStream.closeNode();
      });
      xmlStream.closeNode();
    }
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'dataValidations':
        this.model = {};
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
        list.forEach(addr => {
          if (addr.includes(':')) {
            const range = new Range(addr);
            range.forEachAddress((address: string) => {
              this.model![address] = this._dataValidation!;
            });
          } else {
            this.model![addr] = this._dataValidation!;
          }
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
            formula = utils.excelToDate(parseFloat(formula as string));
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
