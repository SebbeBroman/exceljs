// StringBuilder - a way to keep string memory operations to a minimum
// while building the strings for the xml files
class StringBuilder {
  private _buf!: string[];

  constructor() {
    this.reset();
  }

  get length(): number {
    return this._buf.length;
  }

  toString(): string {
    return this._buf.join('');
  }

  reset(position?: number): void {
    if (position) {
      while (this._buf.length > position) {
        this._buf.pop();
      }
    } else {
      this._buf = [];
    }
  }

  addText(text: string): void {
    this._buf.push(text);
  }

  addStringBuf(inBuf: {toString(): string}): void {
    this._buf.push(inBuf.toString());
  }
}

export default StringBuilder;
export {StringBuilder};
