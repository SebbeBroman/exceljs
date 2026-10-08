import {EventEmitter} from 'node:events';
import {Writable} from 'node:stream';
import {Zip, ZipDeflate, ZipPassThrough, strToU8} from 'fflate';
import {asUint8Array, fromBase64, toPublic} from './bytes.js';

export interface StreamZipWriterOptions {
  level?: number;
  zlib?: {level?: number};
  compression?: string;
  store?: boolean;
  name?: string;
  base64?: boolean;
}

function toUint8Array(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (typeof data === 'string') return strToU8(data);
  if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) return asUint8Array(data);
  return strToU8(String(data));
}

function resolveLevel(options: StreamZipWriterOptions = {}): number {
  if (options.level != null) return options.level;
  if (options.zlib?.level != null) return options.zlib.level;
  if (options.compression === 'STORE' || options.store) return 0;
  return 1;
}

function normalizeName(name: string): string {
  return name.startsWith('/') ? name.slice(1) : name;
}

/** Native writable entries batch XML before compression and wait for output drain. */
class StreamZipWriter extends EventEmitter {
  private readonly zip: Zip;
  private readonly level: number;
  private destination?: Writable;
  private readonly entries = new Set<Writable>();
  private blocked = false;
  private error?: Error;
  private finalized = false;
  private ended = false;
  private readonly waiters = new Set<(error?: Error) => void>();

  constructor(options: StreamZipWriterOptions = {}) {
    super();
    this.level = resolveLevel(options);
    this.zip = new Zip((error, data, final) => {
      if (error) return this.abort(error);
      if (this.error) return;
      try {
        if (data.length && this.destination && !this.destination.write(toPublic(data))) {
          this.blocked = true;
        }
        if (final) this.destination?.end();
      } catch (cause) {
        this.abort(cause as Error);
      }
    });
  }

  abort(error: Error): void {
    if (this.error) return;
    this.error = error;
    for (const settle of this.waiters) settle(error);
    this.waiters.clear();
    for (const stream of this.entries) stream.destroy(error);
    this.destination?.destroy(error);
    this.emit('error', error);
  }

  private createEntry(name: string): ZipDeflate | ZipPassThrough {
    if (this.error) throw this.error;
    if (this.finalized) throw new Error('Cannot append to a finalized archive');
    const entry =
      this.level === 0
        ? new ZipPassThrough(normalizeName(name))
        : new ZipDeflate(normalizeName(name), {
            level: this.level as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9,
          });
    this.zip.add(entry);
    return entry;
  }

  private waitForOutput(callback: (error?: Error) => void): void {
    if (this.error) callback(this.error);
    else if (this.blocked) this.waiters.add(callback);
    else callback();
  }

  private maybeEnd(): void {
    if (this.error || this.ended || !this.finalized || this.entries.size) return;
    this.ended = true;
    try {
      this.zip.end();
    } catch (error) {
      this.abort(error as Error);
    }
  }

  openEntry(name: string): Writable {
    const entry = this.createEntry(name);
    const batch = Buffer.allocUnsafe(65536);
    let length = 0;
    const owner = this;
    const stream = new Writable({
      highWaterMark: batch.length,
      write(chunk: Buffer, _encoding, callback) {
        let offset = 0;
        const pump = (): void => {
          try {
            while (offset < chunk.length) {
              const copied = chunk.copy(batch, length, offset, offset + batch.length - length);
              length += copied;
              offset += copied;
              if (length === batch.length) {
                entry.push(owner.level === 0 ? Buffer.from(batch) : batch, false);
                length = 0;
                if (owner.blocked) {
                  owner.waitForOutput(error => (error ? callback(error) : pump()));
                  return;
                }
              }
            }
            callback();
          } catch (error) {
            callback(error as Error);
          }
        };
        owner.waitForOutput(error => (error ? callback(error) : pump()));
      },
      final(callback) {
        try {
          entry.push(Buffer.from(batch.subarray(0, length)), true);
          owner.waitForOutput(callback);
        } catch (error) {
          callback(error as Error);
        }
      },
    });
    this.entries.add(stream);
    stream.on('error', error => this.abort(error));
    stream.on('finish', () => {
      this.entries.delete(stream);
      this.maybeEnd();
    });
    return stream;
  }

  pipe(destination: Writable): Writable {
    this.destination = destination;
    destination.on('error', error => this.abort(error));
    destination.on('close', () => {
      if (!destination.writableFinished)
        this.abort(new Error('ZIP destination closed before finishing'));
    });
    destination.on('drain', () => {
      this.blocked = false;
      const waiters = [...this.waiters];
      this.waiters.clear();
      for (const settle of waiters) settle();
    });
    return destination;
  }

  append(data: unknown, options: StreamZipWriterOptions = {}): this {
    if (!options.name) throw new Error('zip.append requires options.name');
    const bytes =
      options.base64 && typeof data === 'string' ? fromBase64(data) : toUint8Array(data);
    this.createEntry(options.name).push(bytes, true);
    return this;
  }

  finalize(): void {
    this.finalized = true;
    this.maybeEnd();
  }
}

export default StreamZipWriter;
export {StreamZipWriter, toUint8Array, resolveLevel, normalizeName};
