import {isEqual} from '../../../utils/object.js';
import {dateToExcel} from '../../../utils/utils.js';
import colCache from '../../../utils/col-cache.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
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
            xmlStream.writeText(dateToExcel(new Date(formula as string | number | Date)));
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
}

export default DataValidationsXform;
export {DataValidationsXform};
