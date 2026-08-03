import {Buffer} from 'buffer';
import {strFromU8, unzip} from 'fflate';

function toUint8Array(data) {
  if (data instanceof Uint8Array && !Buffer.isBuffer(data)) {
    return data;
  }
  if (Buffer.isBuffer(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  throw new Error('zip data must be a Buffer, ArrayBuffer, or Uint8Array');
}

/**
 * @param {Buffer|Uint8Array|ArrayBuffer} data
 * @returns {Promise<Record<string, Uint8Array>>}
 */
export function unzipToFiles(data) {
  const u8 = toUint8Array(data);
  return new Promise((resolve, reject) => {
    unzip(u8, (err, files) => {
      if (err) reject(err);
      else resolve(files);
    });
  });
}

/**
 * Decode zip entry bytes as UTF-8 text (XML parts).
 * @param {Uint8Array} u8
 */
export function entryToString(u8) {
  return strFromU8(u8);
}

/**
 * @param {Uint8Array} u8
 */
export function entryToBuffer(u8) {
  return Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength);
}
