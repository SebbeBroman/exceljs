/**
 * Portable byte-buffer helpers (Uint8Array).
 * Replaces the npm `buffer` package so browser bundles stay free of that polyfill.
 *
 * On Node, public-facing results can be wrapped with the native global Buffer
 * (a Uint8Array subclass) for compatibility with Buffer.isBuffer / .toString().
 */

const textEncoder = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;
const textDecoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : null;

type BufferConstructor = typeof globalThis extends {Buffer: infer B}
  ? B extends {from: unknown; isBuffer: unknown}
    ? B
    : typeof Buffer
  : typeof Buffer;

/** @returns Node Buffer constructor or null */
function nativeBuffer(): BufferConstructor | null {
  const B = (globalThis as {Buffer?: BufferConstructor}).Buffer;
  return typeof B === 'function' && typeof (B as typeof Buffer).from === 'function' ? B : null;
}

/**
 * Allocate zeroed bytes.
 */
export function alloc(size: number): Uint8Array {
  return new Uint8Array(size >>> 0);
}

/**
 * True for Uint8Array and Node Buffer (Buffer extends Uint8Array).
 */
export function isBytes(value: unknown): value is Uint8Array {
  return value instanceof Uint8Array;
}

/**
 */
export function concat(
  list: ArrayLike<Uint8Array | null | undefined>,
  totalLength?: number,
): Uint8Array {
  if (!list || list.length === 0) {
    return alloc(0);
  }
  let length = totalLength;
  if (length == null) {
    length = 0;
    for (let i = 0; i < list.length; i++) {
      length += list[i] ? list[i]!.length : 0;
    }
  }
  const out = alloc(length);
  let offset = 0;
  for (let i = 0; i < list.length; i++) {
    const part = list[i];
    if (!part || !part.length) continue;
    const copyLen = Math.min(part.length, length - offset);
    if (copyLen <= 0) break;
    out.set(part.length === copyLen ? part : part.subarray(0, copyLen), offset);
    offset += copyLen;
  }
  return out;
}

/**
 * Copy bytes (Buffer#copy semantics).
 * @returns bytes copied
 */
export function copy(
  src: Uint8Array,
  target: Uint8Array,
  targetStart = 0,
  sourceStart = 0,
  sourceEnd = src.length,
): number {
  const start = sourceStart >>> 0;
  const end = Math.min(sourceEnd >>> 0, src.length);
  const len = Math.max(0, end - start);
  if (len === 0) return 0;
  target.set(src.subarray(start, start + len), targetStart >>> 0);
  return len;
}

/**
 * Decode bytes to string.
 */
export function toString(bytes: Uint8Array, encoding = 'utf8'): string {
  const enc = (encoding || 'utf8').toLowerCase();
  if (enc === 'utf8' || enc === 'utf-8') {
    if (textDecoder) {
      return textDecoder.decode(bytes);
    }
    const B = nativeBuffer();
    if (B)
      return (B as typeof Buffer)
        .from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
        .toString('utf8');
    throw new Error('TextDecoder is required to decode UTF-8');
  }
  if (enc === 'base64') {
    return encodeBase64(bytes);
  }
  if (enc === 'binary' || enc === 'latin1') {
    let s = '';
    for (let i = 0; i < bytes.length; i++) {
      s += String.fromCharCode(bytes[i]);
    }
    return s;
  }
  if (enc === 'hex') {
    let s = '';
    for (let i = 0; i < bytes.length; i++) {
      s += bytes[i].toString(16).padStart(2, '0');
    }
    return s;
  }
  if (enc === 'utf16le' || enc === 'ucs2' || enc === 'ucs-2') {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let s = '';
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      s += String.fromCharCode(view.getUint16(i, true));
    }
    return s;
  }
  // Fall back to utf8
  return toString(bytes, 'utf8');
}

/**
 * Encode string as UTF-8 into buf at offset (Buffer#write-like).
 * Writes as many complete bytes as fit; returns new offset (offset + bytes written).
 */
export function utf8Write(buf: Uint8Array, text: string, offset: number): number {
  const encoded = utf8Encode(text);
  const space = buf.length - offset;
  if (space <= 0) return offset;
  const n = Math.min(encoded.length, space);
  buf.set(encoded.subarray(0, n), offset);
  return offset + n;
}

/**
 */
export function utf8Encode(text: string): Uint8Array {
  if (textEncoder) {
    return textEncoder.encode(text);
  }
  const B = nativeBuffer();
  if (B) {
    const b = (B as typeof Buffer).from(text, 'utf8');
    return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
  }
  throw new Error('TextEncoder is required to encode UTF-8');
}

/**
 */
export function utf16leEncode(text: string): Uint8Array {
  const out = alloc(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = (c >> 8) & 0xff;
  }
  return out;
}

/**
 */
export function fromBase64(b64: string): Uint8Array {
  const B = nativeBuffer();
  if (B) {
    const b = (B as typeof Buffer).from(b64, 'base64');
    return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
  }
  if (typeof atob === 'function') {
    const bin = atob(b64);
    const out = alloc(bin.length);
    for (let i = 0; i < bin.length; i++) {
      out[i] = bin.charCodeAt(i);
    }
    return out;
  }
  throw new Error('No base64 decoder available');
}

/**
 */
export function encodeBase64(bytes: Uint8Array): string {
  const B = nativeBuffer();
  if (B) {
    return (B as typeof Buffer)
      .from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      .toString('base64');
  }
  if (typeof btoa === 'function') {
    const chunk = 0x8000;
    let binary = '';
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }
  throw new Error('No base64 encoder available');
}

/**
 * Buffer.from-compatible constructor.
 */
export function from(
  value: string | ArrayBuffer | ArrayBufferView | ArrayLike<number>,
  encodingOrOffset?: string | number,
  length?: number,
): Uint8Array {
  if (typeof value === 'string') {
    const encoding = ((encodingOrOffset || 'utf8') + '').toLowerCase();
    if (encoding === 'utf8' || encoding === 'utf-8') {
      return utf8Encode(value);
    }
    if (encoding === 'base64') {
      return fromBase64(value);
    }
    if (encoding === 'binary' || encoding === 'latin1') {
      const out = alloc(value.length);
      for (let i = 0; i < value.length; i++) {
        out[i] = value.charCodeAt(i) & 0xff;
      }
      return out;
    }
    if (encoding === 'hex') {
      const hex = value.replace(/^0x/i, '');
      const out = alloc(hex.length >> 1);
      for (let i = 0; i < out.length; i++) {
        out[i] = parseInt(hex.substr(i * 2, 2), 16);
      }
      return out;
    }
    if (encoding === 'utf16le' || encoding === 'ucs2' || encoding === 'ucs-2') {
      return utf16leEncode(value);
    }
    return utf8Encode(value);
  }

  if (value instanceof ArrayBuffer) {
    if (typeof encodingOrOffset === 'number') {
      const start = encodingOrOffset >>> 0;
      const end = length != null ? start + (length >>> 0) : value.byteLength;
      return new Uint8Array(value.slice(start, end));
    }
    return new Uint8Array(value.slice(0));
  }

  if (ArrayBuffer.isView(value)) {
    // Copy so callers never share fflate / pool memory unexpectedly
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength).slice();
  }

  if (Array.isArray(value) || (value && typeof (value as ArrayLike<number>).length === 'number')) {
    return Uint8Array.from(value as ArrayLike<number>);
  }

  throw new TypeError('bytes.from: unsupported value type');
}

/**
 * Normalize any accepted binary input to a Uint8Array (copy when needed for safety).
 * Accepts Buffer, Uint8Array, ArrayBuffer, ArrayBufferView.
 */
export function asUint8Array(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) {
    // Share view for Uint8Array/Buffer — callers that need isolation should copy
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  throw new TypeError('zip data must be a Buffer, ArrayBuffer, or Uint8Array');
}

/**
 * Wrap for public API: Node Buffer when available, else Uint8Array.
 * Zero-copy when Node Buffer can view the same memory.
 */
export function toPublic(bytes: Uint8Array | null | undefined): Uint8Array | null | undefined {
  if (bytes == null) return bytes;
  const B = nativeBuffer();
  if (B) {
    if ((B as typeof Buffer).isBuffer(bytes)) return bytes;
    return (B as typeof Buffer).from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  return bytes;
}

/**
 * Write uint32 little-endian at offset.
 */
export function writeUInt32LE(buf: Uint8Array, value: number, offset = 0): void {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  view.setUint32(offset, value >>> 0, true);
}
