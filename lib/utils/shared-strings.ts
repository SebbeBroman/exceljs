class SharedStrings {
  private _values: string[];
  private _totalRefs: number;
  private _hash: Record<string, number>;

  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null) as Record<string, number>;
  }

  get count(): number {
    return this._values.length;
  }

  get values(): string[] {
    return this._values;
  }

  get totalRefs(): number {
    return this._totalRefs;
  }

  getString(index: number): string | undefined {
    return this._values[index];
  }

  add(value: string): number {
    let index = this._hash[value];
    if (index === undefined) {
      index = this._hash[value] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

export default SharedStrings;
export {SharedStrings};
