import {EventEmitter} from './event-emitter.js';
import fs from 'fs';
import {Writable} from 'stream';
import {Zip, ZipDeflate, ZipPassThrough, strToU8} from 'fflate';
import {asUint8Array, fromBase64, isBytes, toPublic} from './bytes.js';

function toUint8Array(data) {
  if (typeof data === 'string') {
    return strToU8(data);
  }
  if (isBytes(data) || data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
    return asUint8Array(data);
  }
  return strToU8(String(data));
}

function resolveLevel(options = {}) {
  if (options.level != null) return options.level;
  if (options.zlib && options.zlib.level != null) return options.zlib.level;
  // JSZip / archiver used compression: 'DEFLATE' | 'STORE'
  if (options.compression === 'STORE') return 0;
  if (options.store === true) return 0;
  return 6;
}

function normalizeName(name) {
  if (!name) return name;
  return name.charAt(0) === '/' ? name.slice(1) : name;
}

function isStreamLike(data) {
  if (data == null || typeof data !== 'object') return false;
  if (isBytes(data)) return false;
  if (data instanceof ArrayBuffer) return false;
  if (ArrayBuffer.isView(data)) return false;
  // Readable / StreamBuf / EventEmitter-based sources
  return typeof data.on === 'function';
}

/**
 * Streaming zip writer backed by fflate's Zip / ZipDeflate / ZipPassThrough.
 * Archiver-like surface used by the streaming xlsx WorkbookWriter:
 *   pipe(dest), append(data|stream, {name, base64?}), file(fsPath, {name}), finalize(), 'error'
 *
 * StreamBuf note: worksheet streams are pause()d so 'data' is not emitted; archiver
 * consumed them via pipe(). StreamBuf's custom pipe() pushes to destination.write
 * regardless of pause, so we always pipe into a Writable sink.
 */
class StreamZipWriter extends EventEmitter {
  constructor(options) {
    super();
    this.options = options || {};
    this.level = resolveLevel(this.options);
    this.destination = null;
    this.pending = 0;
    this.finalized = false;
    this.ended = false;
    this.errored = false;

    this._zip = new Zip((err, data, final) => {
      if (err) {
        this._fail(err);
        return;
      }
      if (data && data.length && this.destination) {
        // Best-effort write; do not drop chunks
        this.destination.write(toPublic(data));
      }
      if (final && this.destination) {
        this.destination.end();
      }
    });
  }

  _fail(err) {
    if (this.errored) return;
    this.errored = true;
    this.emit('error', err);
  }

  _createEntry(name) {
    const filename = normalizeName(name);
    let entry;
    if (this.level === 0) {
      entry = new ZipPassThrough(filename);
    } else {
      entry = new ZipDeflate(filename, {level: this.level});
    }
    this._zip.add(entry);
    return entry;
  }

  _beginEntry() {
    this.pending += 1;
    let settled = false;
    return () => {
      if (settled) return;
      settled = true;
      this.pending -= 1;
      this._maybeEnd();
    };
  }

  _maybeEnd() {
    if (this.errored || this.ended || !this.finalized || this.pending > 0) {
      return;
    }
    this.ended = true;
    try {
      this._zip.end();
    } catch (err) {
      this._fail(err);
    }
  }

  /**
   * Writable sink that pushes source chunks into an fflate zip entry.
   * Works with both Node Readable.pipe and StreamBuf's custom pipe().
   */
  _createEntrySink(entry, done) {
    let finished = false;
    const self = this;

    const complete = err => {
      if (finished) return;
      finished = true;
      if (err) {
        self._fail(err);
        done();
        return;
      }
      try {
        entry.push(new Uint8Array(0), true);
      } catch (e) {
        self._fail(e);
      }
      done();
    };

    const sink = new Writable({
      write(chunk, encoding, callback) {
        try {
          entry.push(toUint8Array(chunk), false);
          callback();
        } catch (err) {
          complete(err);
          callback(err);
        }
      },
      final(callback) {
        complete();
        callback();
      },
      destroy(err, callback) {
        if (err) complete(err);
        callback(err);
      },
    });

    sink.on('error', err => complete(err));
    return sink;
  }

  pipe(destination) {
    this.destination = destination;
    if (destination && typeof destination.on === 'function') {
      destination.on('error', err => this._fail(err));
    }
    return destination;
  }

  /**
   * Append a string/Buffer/Uint8Array or a Readable-like stream as a zip entry.
   */
  append(data, options = {}) {
    const name = options.name;
    if (!name) {
      throw new Error('zip.append requires options.name');
    }

    if (isStreamLike(data)) {
      const entry = this._createEntry(name);
      const done = this._beginEntry();
      const sink = this._createEntrySink(entry, done);

      data.on('error', err => {
        this._fail(err);
        // ensure pending is released if sink never gets final
        if (typeof sink.destroy === 'function') {
          sink.destroy(err);
        } else {
          done();
        }
      });

      if (typeof data.pipe === 'function') {
        // StreamBuf: custom pipe uses destination.write / .end
        // Node Readable: standard pipe into Writable
        data.pipe(sink);
      } else {
        // Fallback: EventEmitter readable protocol
        let finished = false;
        const finish = () => {
          if (finished) return;
          finished = true;
          sink.end();
        };
        data.on('data', chunk => {
          sink.write(chunk);
        });
        data.on('end', finish);
        data.on('finish', finish);
      }
      return this;
    }

    // Static payload
    const entry = this._createEntry(name);
    let bytes;
    if (options.base64) {
      bytes = typeof data === 'string' ? fromBase64(data) : toUint8Array(data);
    } else {
      bytes = toUint8Array(data);
    }
    entry.push(bytes, true);
    return this;
  }

  /**
   * Read a file from disk into a zip entry (streaming).
   * Returns a Promise that resolves when the entry has been fully pushed
   * (so Promise.all in addMedia can wait if desired).
   */
  file(fsPath, options = {}) {
    const name = options.name;
    if (!name) {
      return Promise.reject(new Error('zip.file requires options.name'));
    }

    return new Promise((resolve, reject) => {
      const entry = this._createEntry(name);
      const done = this._beginEntry();
      const stream = fs.createReadStream(fsPath);
      let settled = false;

      const settle = err => {
        if (settled) return;
        settled = true;
        if (err) {
          this._fail(err);
          done();
          reject(err);
          return;
        }
        done();
        resolve();
      };

      const sink = new Writable({
        write(chunk, encoding, callback) {
          try {
            entry.push(toUint8Array(chunk), false);
            callback();
          } catch (e) {
            callback(e);
          }
        },
        final(callback) {
          try {
            entry.push(new Uint8Array(0), true);
            callback();
          } catch (e) {
            callback(e);
          }
        },
      });

      sink.on('finish', () => settle());
      sink.on('error', err => settle(err));
      stream.on('error', err => settle(err));
      stream.pipe(sink);
    });
  }

  /**
   * Finish the archive. Waits for any in-flight stream/file entries, then
   * ends the fflate Zip (which writes the central directory and closes dest).
   */
  finalize() {
    this.finalized = true;
    this._maybeEnd();
  }
}

export default StreamZipWriter;
export {StreamZipWriter, toUint8Array, resolveLevel, normalizeName};
