import {Buffer} from 'buffer';
import {EventEmitter} from 'events';
import {strToU8, zip} from 'fflate';
import StreamBuf from './stream-buf.js';

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
  if (typeof data === 'string') {
    return strToU8(data);
  }
  return strToU8(String(data));
}

function resolveLevel(options = {}) {
  if (options.level != null) return options.level;
  if (options.zlib && options.zlib.level != null) return options.zlib.level;
  // JSZip used compression: 'DEFLATE' | 'STORE'
  if (options.compression === 'STORE') return 0;
  return 6;
}

function zipAsync(files, level) {
  return new Promise((resolve, reject) => {
    zip(files, {level}, (err, data) => {
      if (err) reject(err);
      else resolve(data);
    });
  });
}

// =============================================================================
// The ZipWriter class
// Packs members into a zip (via fflate), then exposes them as a readable StreamBuf
class ZipWriter extends EventEmitter {
  constructor(options) {
    super();
    this.options = options || {};
    this.level = resolveLevel(this.options);
    /** @type {Record<string, Uint8Array>} */
    this.files = Object.create(null);
    this.stream = new StreamBuf();
  }

  append(data, options) {
    const name = options.name;
    if (!name) {
      throw new Error('zip.append requires options.name');
    }
    let bytes;
    if (options.base64) {
      bytes = toUint8Array(Buffer.from(data, 'base64'));
    } else {
      bytes = toUint8Array(data);
    }
    this.files[name] = bytes;
  }

  async finalize() {
    const content = await zipAsync(this.files, this.level);
    this.stream.end(Buffer.from(content.buffer, content.byteOffset, content.byteLength));
    this.emit('finish');
  }

  // ==========================================================================
  // Stream.Readable interface (used when piping ZipWriter into StreamBuf / files)
  read(size) {
    return this.stream.read(size);
  }

  setEncoding(encoding) {
    return this.stream.setEncoding(encoding);
  }

  pause() {
    return this.stream.pause();
  }

  resume() {
    return this.stream.resume();
  }

  isPaused() {
    return this.stream.isPaused();
  }

  pipe(destination, options) {
    return this.stream.pipe(destination, options);
  }

  unpipe(destination) {
    return this.stream.unpipe(destination);
  }

  unshift(chunk) {
    return this.stream.unshift(chunk);
  }

  wrap(stream) {
    return this.stream.wrap(stream);
  }
}

export default {ZipWriter};
export {ZipWriter};
