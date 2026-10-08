import {isEqual} from '../../../utils/object.js';
import utils from '../../../utils/utils.js';
import colCache from '../../../utils/col-cache.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import Range from '../../../model/range.js';
import {subtractRectangle, mergeRectangles} from '../../../model/rectangles.js';
import type {Rectangle} from '../../../model/rectangles.js';

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
  // Later entries overwrite overlapping earlier rules, matching the old per-cell map.
  const source = Object.entries(model).map(([address, rule]) => {
    if (address.includes(':') || address.includes('!'))
      return {rect: new Range(address).model, rule};
    const {row, col} = colCache.decodeAddress(address);
    return {rect: {top: row!, bottom: row!, left: col!, right: col!}, rule};
  });
  let entries: Array<{rect: Rectangle; rule: DataValidationEntry}> = [];
  if (source.every(({rect}) => rect.top === rect.bottom && rect.left === rect.right)) {
    entries = Object.keys(model).every(address => /^[A-Z]+[1-9]\d*$/.test(address))
      ? source
      : [...new Map(source.map(entry => [`${entry.rect.top}:${entry.rect.left}`, entry])).values()];
  } else
    for (const {rect, rule} of source) {
      entries = entries.flatMap(entry =>
        subtractRectangle(entry.rect, rect).map(piece => ({rect: piece, rule: entry.rule})),
      );
      entries.push({rect, rule});
    }
  const groups: Array<{rects: Rectangle[]; rule: DataValidationEntry}> = [];
  const buckets = new Map<string, typeof groups>();
  const identities = new WeakMap<DataValidationEntry, (typeof groups)[number]>();
  for (const {rect, rule} of entries) {
    const known = identities.get(rule);
    if (known) {
      known.rects.push(rect);
      continue;
    }
    const key = JSON.stringify(
      Object.keys(rule)
        .sort()
        .map(name => [name, rule[name]]),
    );
    const bucket = buckets.get(key) ?? [];
    let group = bucket.find(g => isEqual(g.rule, rule));
    if (!group) {
      group = {rects: [], rule};
      groups.push(group);
      bucket.push(group);
      buckets.set(key, bucket);
    }
    identities.set(rule, group);
    group.rects.push(rect);
  }
  return groups.flatMap(({rects, rule}) =>
    mergeRectangles(rects).map(rect => ({...rule, sqref: new Range(rect).shortRange})),
  );
}

class DataValidationsXform extends BaseXform<DataValidationsModel> {
  _address?: string;
  _dataValidation?: DataValidationEntry;
  _formula?: string[];
  private rangeKeys = new Set<string>();

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
