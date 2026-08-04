import {utf8Encode} from './bytes.js';

function stringToBuffer(str: string | Uint8Array): string | Uint8Array {
  if (typeof str !== 'string') {
    return str;
  }
  return utf8Encode(str);
}
export {stringToBuffer};
