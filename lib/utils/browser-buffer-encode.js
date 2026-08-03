import {utf8Encode} from './bytes.js';

function stringToBuffer(str) {
  if (typeof str !== 'string') {
    return str;
  }
  return utf8Encode(str);
}
export {stringToBuffer};
