import {strFromU8, unzip} from 'fflate';
import {asUint8Array, toPublic} from './bytes.js';

/**
 * @param {Uint8Array|ArrayBuffer|ArrayBufferView} data
 * @returns {Uint8Array}
 */
function toUint8Array(data) {
  return asUint8Array(data);
}

/**
 * @param {Uint8Array|ArrayBuffer|ArrayBufferView} data
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
 * @returns {Uint8Array} public buffer (Node Buffer when available)
 */
export function entryToBuffer(u8) {
  return toPublic(u8 instanceof Uint8Array ? u8 : asUint8Array(u8));
}
