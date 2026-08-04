import {alloc, copy, toPublic, utf8Encode} from './bytes.js';

export interface StringBufOptions {
  size?: number;
  encoding?: string;
}

// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files
class StringBuf {
  _buf: Uint8Array;
  _encoding: string;
  _inPos: number;
  _buffer: Uint8Array | undefined;

  constructor(options?: StringBufOptions) {
    this._buf = alloc((options && options.size) || 16384);
    this._encoding = (options && options.encoding) || 'utf8';

    // where in the buffer we are at
    this._inPos = 0;

    // for use by toBuffer()
    this._buffer = undefined;
  }

  get length(): number {
    return this._inPos;
  }

  get capacity(): number {
    return this._buf.length;
  }

  get buffer(): Uint8Array {
    return this._buf;
  }

  toBuffer(): Uint8Array {
    // return the current data as a single enclosing buffer
    if (!this._buffer) {
      const slice = alloc(this.length);
      copy(this._buf, slice, 0, 0, this.length);
      this._buffer = toPublic(slice) as Uint8Array;
    }
    return this._buffer;
  }

  reset(position?: number): void {
    position = position || 0;
    this._buffer = undefined;
    this._inPos = position;
  }

  _grow(min: number): void {
    let size = this._buf.length * 2;
    while (size < min) {
      size *= 2;
    }
    const buf = alloc(size);
    copy(this._buf, buf, 0, 0, this._inPos);
    this._buf = buf;
  }

  addText(text: string): void {
    this._buffer = undefined;
    const encoded = utf8Encode(text);
    // Grow if write would not fit, or leave <4 bytes free (legacy Buffer#write headroom)
    if (this._inPos + encoded.length > this._buf.length - 4) {
      this._grow(this._inPos + encoded.length + 4);
    }
    this._buf.set(encoded, this._inPos);
    this._inPos += encoded.length;
  }

  addStringBuf(inBuf: StringBuf): void {
    if (inBuf.length) {
      this._buffer = undefined;

      if (this.length + inBuf.length > this.capacity) {
        this._grow(this.length + inBuf.length);
      }

      copy(inBuf._buf, this._buf, this._inPos, 0, inBuf.length);
      this._inPos += inBuf.length;
    }
  }
}

export default StringBuf;
export {StringBuf};
