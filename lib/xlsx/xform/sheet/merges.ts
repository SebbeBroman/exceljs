import _ from '../../../utils/under-dash.js';
import Range from '../../../doc/range.js';
import colCache from '../../../utils/col-cache.js';
import Enums from '../../../doc/enums.js';

export interface MergeInput {
  address: string;
  master: string;
}

export interface MergeCellLike {
  type: number;
  address?: string;
  master?: string;
}

export interface MergeRowLike {
  cells: Array<MergeCellLike | undefined | null>;
}

class Merges {
  merges: Record<string, Range>;
  // used by getMasterAddress after reconcile (may be set externally)
  hash?: Record<string, Range>;

  constructor() {
    // optional mergeCells is array of ranges (like the xml)
    this.merges = {};
  }

  add(merge: MergeInput): void {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      this.merges[merge.master].expandToAddress(merge.address);
    } else {
      const range = `${merge.master}:${merge.address}`;
      this.merges[merge.master] = new Range(range);
    }
  }

  get mergeCells(): string[] {
    return _.map(this.merges, (merge: Range) => merge.range);
  }

  reconcile(mergeCells: string[], rows: MergeRowLike[]): void {
    // reconcile merge list with merge cells
    _.each(mergeCells, (merge: string) => {
      const dimensions = colCache.decode(merge) as {
        top: number;
        left: number;
        bottom: number;
        right: number;
        tl: string;
      };
      for (let i = dimensions.top; i <= dimensions.bottom; i++) {
        const row = rows[i - 1];
        for (let j = dimensions.left; j <= dimensions.right; j++) {
          const cell = row.cells[j - 1];
          if (!cell) {
            // nulls are not included in document - so if master cell has no value - add a null one here
            // Note: original uses j (1-based) as index — preserved as-is
            row.cells[j] = {
              type: Enums.ValueType.Null,
              address: colCache.encodeAddress(i, j),
            };
          } else if (cell.type === Enums.ValueType.Merge) {
            cell.master = dimensions.tl;
          }
        }
      }
    });
  }

  getMasterAddress(address: string): string | undefined {
    // if address has been merged, return its master's address. Assumes reconcile has been called
    const range = this.hash?.[address];
    return range && range.tl;
  }
}

export default Merges;
export {Merges};
