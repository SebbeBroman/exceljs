import Range from './range.js';
import colCache from '../utils/col-cache.js';
import type {DefinedNamesModel} from './schema.js';
import {unionRectangles} from './rectangles.js';
import type {Rectangle} from './rectangles.js';

/** Named references are unions of rectangles, rather than a matrix of cell objects. */
class DefinedNames {
  private names = new Map<string, Map<string | undefined, Rectangle[]>>();

  add(reference: string, name: string): void {
    const decoded = colCache.decodeEx(reference);
    let rect: Rectangle;
    if ('top' in decoded) rect = decoded;
    else if ('row' in decoded && decoded.row != null && decoded.col != null) {
      rect = {top: decoded.row, bottom: decoded.row, left: decoded.col, right: decoded.col};
    } else throw new Error(`Invalid defined name range: ${reference}`);
    const {top, left, bottom, right} = rect;
    const sheetName = 'sheetName' in decoded ? decoded.sheetName : undefined;
    if (
      ![top, left, bottom, right].every(Number.isInteger) ||
      top < 1 ||
      left < 1 ||
      bottom > 1048576 ||
      right > 16384
    )
      throw new Error(`Invalid defined name range: ${reference}`);
    let sheets = this.names.get(name);
    if (!sheets) this.names.set(name, (sheets = new Map()));
    const ranges = sheets.get(sheetName) ?? [];
    ranges.push({top, left, bottom, right});
    sheets.set(sheetName, ranges);
  }

  get model(): DefinedNamesModel {
    return [...this.names].flatMap(([name, sheets]) => {
      const ranges = [...sheets].flatMap(([sheetName, rectangles]) =>
        unionRectangles(rectangles).map(rect => new Range({...rect, sheetName}).$shortRange),
      );
      return ranges.length ? [{name, ranges}] : [];
    });
  }

  set model(value: DefinedNamesModel) {
    this.names.clear();
    for (const entry of value ?? [])
      for (const reference of entry.ranges) {
        // Match the model loader's existing treatment of invalid/non-cell names.
        if (!/[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/.test(reference.split('!').pop() ?? ''))
          continue;
        try {
          this.add(reference, entry.name);
        } catch {
          /* Ignore malformed references from files. */
        }
      }
  }
}
export default DefinedNames;
export {DefinedNames};
