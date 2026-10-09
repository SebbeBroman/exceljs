import Range from '../../../model/range.js';

export interface MergeInput {
  address: string;
  master: string;
}

class Merges {
  merges: Record<string, Range>;

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
    return Object.values(this.merges).map((merge: Range) => merge.range);
  }
}

export default Merges;
export {Merges};
