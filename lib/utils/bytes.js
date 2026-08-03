/**
 * Portable byte-buffer helpers (Uint8Array).
 * Replaces the npm `buffer` package so browser bundles stay free of that polyfill.
 *
 * On Node, public-facing results can be wrapped with the native global Buffer
 * (a Uint8Array subclass) for compatibility with Buffer.isBuffer / .toString().
 */

const textEncoder = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;
const textDecoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : null;

/** @returns {typeof globalThis.Buffer | null} */
function nativeBuffer() {
  const B = globalThis.Buffer;
  return typeof B === 'function' && typeof B.from === 'function' ? B : null;
}

/**
 * Allocate zeroed bytes.
 * @param {number} size
 * @returns {Uint8Array}
 */
export function alloc(size) {
  return new Uint8Array(size >>> 0);
}

/**
 * True for Uint8Array and Node Buffer (Buffer extends Uint8Array).
 * @param {unknown} value
 * @returns {value is Uint8Array}
 */
export function isBytes(value) {
  return value instanceof Uint8Array;
}

/**
 * @param {Uint8Array[]} list
 * @param {number} [totalLength]
 * @returns {Uint8Array}
 */
export function concat(list, totalLength) {
  if (!list || list.length === 0) {
    return alloc(0);
  }
  let length = totalLength;
  if (length == null) {
    length = 0;
    for (let i = 0; i < list.length; i++) {
      length += list[i] ? list[i].length : 0;
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
 * @param {Uint8Array} src
 * @param {Uint8Array} target
 * @param {number} [targetStart]
 * @param {number} [sourceStart]
 * @param {number} [sourceEnd]
 * @returns {number} bytes copied
 */
export function copy(src, target, targetStart = 0, sourceStart = 0, sourceEnd = src.length) {
  const start = sourceStart >>> 0;
  const end = Math.min(sourceEnd >>> 0, src.length);
  const len = Math.max(0, end - start);
  if (len === 0) return 0;
  target.set(src.subarray(start, start + len), targetStart >>> 0);
  return len;
}

/**
 * Decode bytes to string.
 * @param {Uint8Array} bytes
 * @param {string} [encoding]
 * @returns {string}
 */
export function toString(bytes, encoding = 'utf8') {
  const enc = (encoding || 'utf8').toLowerCase();
  if (enc === 'utf8' || enc === 'utf-8') {
    if (textDecoder) {
      return textDecoder.decode(bytes);
    }
    const B = nativeBuffer();
    if (B) return B.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('utf8');
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
 * @param {Uint8Array} buf
 * @param {string} text
 * @param {number} offset
 * @returns {number}
 */
export function utf8Write(buf, text, offset) {
  const encoded = utf8Encode(text);
  const space = buf.length - offset;
  if (space <= 0) return offset;
  const n = Math.min(encoded.length, space);
  buf.set(encoded.subarray(0, n), offset);
  return offset + n;
}

/**
 * @param {string} text
 * @returns {Uint8Array}
 */
export function utf8Encode(text) {
  if (textEncoder) {
    return textEncoder.encode(text);
  }
  const B = nativeBuffer();
  if (B) {
    const b = B.from(text, 'utf8');
    return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
  }
  throw new Error('TextEncoder is required to encode UTF-8');
}

/**
 * @param {string} text
 * @returns {Uint8Array}
 */
export function utf16leEncode(text) {
  const out = alloc(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = (c >> 8) & 0xff;
  }
  return out;
}

/**
 * @param {string} b64
 * @returns {Uint8Array}
 */
export function fromBase64(b64) {
  const B = nativeBuffer();
  if (B) {
    const b = B.from(b64, 'base64');
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
 * @param {Uint8Array} bytes
 * @returns {string}
 */
export function encodeBase64(bytes) {
  const B = nativeBuffer();
  if (B) {
    return B.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64');
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
 * @param {string|ArrayBuffer|ArrayBufferView|ArrayLike<number>} value
 * @param {string|number} [encodingOrOffset]
 * @param {number} [length]
 * @returns {Uint8Array}
 */
export function from(value, encodingOrOffset, length) {
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

  if (Array.isArray(value) || (value && typeof value.length === 'number')) {
    return Uint8Array.from(value);
  }

  throw new TypeError('bytes.from: unsupported value type');
}

/**
 * Normalize any accepted binary input to a Uint8Array (copy when needed for safety).
 * Accepts Buffer, Uint8Array, ArrayBuffer, ArrayBufferView.
 * @param {unknown} data
 * @returns {Uint8Array}
 */
export function asUint8Array(data) {
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
 * @param {Uint8Array|null|undefined} bytes
 * @returns {Uint8Array|null|undefined}
 */
export function toPublic(bytes) {
  if (bytes == null) return bytes;
  const B = nativeBuffer();
  if (B) {
    if (B.isBuffer(bytes)) return bytes;
    return B.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  return bytes;
}

/**
 * Write uint32 little-endian at offset.
 * @param {Uint8Array} buf
 * @param {number} value
 * @param {number} offset
 */
export function writeUInt32LE(buf, value, offset = 0) {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  view.setUint32(offset, value >>> 0, true);
}
