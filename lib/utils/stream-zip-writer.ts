import {EventEmitter} from './event-emitter.js';
import fs from 'fs';
import {Writable} from 'stream';
import {Zip, ZipDeflate, ZipPassThrough, strToU8} from 'fflate';
import type {
  Zip as ZipType,
  ZipDeflate as ZipDeflateType,
  ZipPassThrough as ZipPassThroughType,
} from 'fflate';
import {asUint8Array, fromBase64, isBytes, toPublic} from './bytes.js';

export interface StreamZipWriterOptions {
  level?: number;
  zlib?: {level?: number};
  compression?: string;
  store?: boolean;
  name?: string;
  base64?: boolean;
}

type ZipEntry = ZipDeflateType | ZipPassThroughType;

/** Stream-like source with event / pipe surface. */
interface StreamLike {
  on(event: string, listener: (...args: unknown[]) => void): unknown;
  pipe?(destination: Writable): unknown;
}

/** Destination writable for zip output. */
interface ZipDestination {
  write(chunk: unknown): unknown;
  end(): unknown;
  on?(event: string, listener: (...args: unknown[]) => void): unknown;
}

function toUint8Array(data: unknown): Uint8Array {
  // Already bytes — share the view (no re-copy / re-encode)
  if (data instanceof Uint8Array) {
    return data;
  }
  if (typeof data === 'string') {
    return strToU8(data);
  }
  if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
    return asUint8Array(data);
  }
  return strToU8(String(data));
}

function resolveLevel(options: StreamZipWriterOptions = {}): number {
  if (options.level != null) return options.level;
  if (options.zlib && options.zlib.level != null) return options.zlib.level;
  // JSZip / archiver used compression: 'DEFLATE' | 'STORE'
  if (options.compression === 'STORE') return 0;
  if (options.store === true) return 0;
  // Default level 1: much faster than 6 with similar size order-of-magnitude
  return 1;
}

function normalizeName(name: string | undefined): string | undefined {
  if (!name) return name;
  return name.charAt(0) === '/' ? name.slice(1) : name;
}

function isStreamLike(data: unknown): data is StreamLike {
  if (data == null || typeof data !== 'object') return false;
  if (isBytes(data)) return false;
  if (data instanceof ArrayBuffer) return false;
  if (ArrayBuffer.isView(data)) return false;
  // Readable / StreamBuf / EventEmitter-based sources
  return typeof (data as StreamLike).on === 'function';
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
  options: StreamZipWriterOptions;
  level: number;
  destination: ZipDestination | null;
  pending: number;
  finalized: boolean;
  ended: boolean;
  errored: boolean;
  _zip: ZipType;

  constructor(options?: StreamZipWriterOptions) {
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

  _fail(err: Error): void {
    if (this.errored) return;
    this.errored = true;
    this.emit('error', err);
  }

  _createEntry(name: string): ZipEntry {
    const filename = normalizeName(name)!;
    let entry: ZipEntry;
    if (this.level === 0) {
      entry = new ZipPassThrough(filename);
    } else {
      entry = new ZipDeflate(filename, {
        level: this.level as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9,
      });
    }
    this._zip.add(entry);
    return entry;
  }

  _beginEntry(): () => void {
    this.pending += 1;
    let settled = false;
    return () => {
      if (settled) return;
      settled = true;
      this.pending -= 1;
      this._maybeEnd();
    };
  }

  _maybeEnd(): void {
    if (this.errored || this.ended || !this.finalized || this.pending > 0) {
      return;
    }
    this.ended = true;
    try {
      this._zip.end();
    } catch (err) {
      this._fail(err as Error);
    }
  }

  /**
   * Writable sink that pushes source chunks into an fflate zip entry.
   * Works with both Node Readable.pipe and StreamBuf's custom pipe().
   */
  _createEntrySink(entry: ZipEntry, done: () => void): Writable {
    let finished = false;
    const self = this;

    const complete = (err?: Error | null): void => {
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
        self._fail(e as Error);
      }
      done();
    };

    const sink = new Writable({
      write(chunk, _encoding, callback) {
        try {
          entry.push(toUint8Array(chunk), false);
          callback();
        } catch (err) {
          complete(err as Error);
          callback(err as Error);
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

  pipe(destination: ZipDestination): ZipDestination {
    this.destination = destination;
    if (destination && typeof destination.on === 'function') {
      destination.on('error', (err: unknown) => this._fail(err as Error));
    }
    return destination;
  }

  /**
   * Append a string/Buffer/Uint8Array or a Readable-like stream as a zip entry.
   */
  append(data: unknown, options: StreamZipWriterOptions = {}): this {
    const name = options.name;
    if (!name) {
      throw new Error('zip.append requires options.name');
    }

    if (isStreamLike(data)) {
      const entry = this._createEntry(name);
      const done = this._beginEntry();
      const sink = this._createEntrySink(entry, done);

      data.on('error', (err: unknown) => {
        this._fail(err as Error);
        // ensure pending is released if sink never gets final
        if (typeof sink.destroy === 'function') {
          sink.destroy(err as Error);
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
        const finish = (): void => {
          if (finished) return;
          finished = true;
          sink.end();
        };
        data.on('data', (chunk: unknown) => {
          sink.write(chunk as Buffer);
        });
        data.on('end', finish);
        data.on('finish', finish);
      }
      return this;
    }

    // Static payload
    const entry = this._createEntry(name);
    let bytes: Uint8Array;
    if (options.base64) {
      bytes = typeof data === 'string' ? fromBase64(data) : toUint8Array(data);
    } else if (data instanceof Uint8Array) {
      // Zero-copy: already encoded bytes
      bytes = data;
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
  file(fsPath: string, options: StreamZipWriterOptions = {}): Promise<void> {
    const name = options.name;
    if (!name) {
      return Promise.reject(new Error('zip.file requires options.name'));
    }

    return new Promise((resolve, reject) => {
      const entry = this._createEntry(name);
      const done = this._beginEntry();
      const stream = fs.createReadStream(fsPath);
      let settled = false;

      const settle = (err?: Error | null): void => {
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
        write(chunk, _encoding, callback) {
          try {
            entry.push(toUint8Array(chunk), false);
            callback();
          } catch (e) {
            callback(e as Error);
          }
        },
        final(callback) {
          try {
            entry.push(new Uint8Array(0), true);
            callback();
          } catch (e) {
            callback(e as Error);
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
  finalize(): void {
    this.finalized = true;
    this._maybeEnd();
  }
}

export default StreamZipWriter;
export {StreamZipWriter, toUint8Array, resolveLevel, normalizeName};
