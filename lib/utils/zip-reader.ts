import {strFromU8, unzip} from 'fflate';
import {asUint8Array, toPublic} from './bytes.js';

/**
 * @param data
 * @returns
 */
function toUint8Array(data: ArrayBuffer | ArrayBufferView): Uint8Array {
  return asUint8Array(data);
}

/**
 * @param data
 * @returns
 */
export function unzipToFiles(
  data: ArrayBuffer | ArrayBufferView,
): Promise<Record<string, Uint8Array>> {
  const u8 = toUint8Array(data);
  return new Promise((resolve, reject) => {
    unzip(u8, (err, files) => {
      if (err) reject(err);
      else resolve(files as Record<string, Uint8Array>);
    });
  });
}

/**
 * Decode zip entry bytes as UTF-8 text (XML parts).
 * @param u8
 */
export function entryToString(u8: Uint8Array): string {
  return strFromU8(u8);
}

/**
 * @param u8
 * @returns public buffer (Node Buffer when available)
 */
export function entryToBuffer(u8: Uint8Array | ArrayBufferView): Uint8Array {
  return toPublic(u8 instanceof Uint8Array ? u8 : asUint8Array(u8)) as Uint8Array;
}
