import {toString as bytesToString} from './bytes.js';

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (chunk instanceof Uint8Array || ArrayBuffer.isView(chunk)) {
    return bytesToString(
      chunk instanceof Uint8Array
        ? chunk
        : new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength)
    );
  }
  if (chunk && typeof chunk.toString === 'function') {
    return chunk.toString();
  }
  return String(chunk);
}
export {bufferToString};
