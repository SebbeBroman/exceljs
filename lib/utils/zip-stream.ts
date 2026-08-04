import {EventEmitter} from './event-emitter.js';
import {strToU8, zip} from 'fflate';
import StreamBuf from './stream-buf.js';
import {asUint8Array, fromBase64, toPublic} from './bytes.js';

export interface ZipOptions {
  level?: number;
  zlib?: {level?: number};
  compression?: string;
  name?: string;
  base64?: boolean;
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

function resolveLevel(options: ZipOptions = {}): number {
  if (options.level != null) return options.level;
  if (options.zlib && options.zlib.level != null) return options.zlib.level;
  // JSZip used compression: 'DEFLATE' | 'STORE'
  if (options.compression === 'STORE') return 0;
  // Default level 1: much faster than 6 with similar size order-of-magnitude
  return 1;
}

type ZipLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

function zipAsync(
  files: Record<string, Uint8Array>,
  level: number,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    zip(files, {level: level as ZipLevel}, (err, data) => {
      if (err) reject(err);
      else resolve(data);
    });
  });
}

// =============================================================================
// The ZipWriter class
// Packs members into a zip (via fflate), then exposes them as a readable StreamBuf
class ZipWriter extends EventEmitter {
  options: ZipOptions;
  level: number;
  files: Record<string, Uint8Array>;
  stream: InstanceType<typeof StreamBuf>;

  constructor(options?: ZipOptions) {
    super();
    this.options = options || {};
    this.level = resolveLevel(this.options);
    this.files = Object.create(null) as Record<string, Uint8Array>;
    this.stream = new StreamBuf();
  }

  append(data: unknown, options: ZipOptions): void {
    const name = options.name;
    if (!name) {
      throw new Error('zip.append requires options.name');
    }
    let bytes: Uint8Array;
    if (options.base64) {
      bytes = typeof data === 'string' ? fromBase64(data) : toUint8Array(data);
    } else if (data instanceof Uint8Array) {
      // Zero-copy: already encoded bytes
      bytes = data;
    } else {
      bytes = toUint8Array(data);
    }
    this.files[name] = bytes;
  }

  async finalize(): Promise<void> {
    const content = await zipAsync(this.files, this.level);
    this.stream.end(toPublic(content));
    this.emit('finish');
  }

  // ==========================================================================
  // Stream.Readable interface (used when piping ZipWriter into StreamBuf / files)
  read(size?: number): Uint8Array | null | undefined {
    return this.stream.read(size);
  }

  setEncoding(encoding: string): void {
    return this.stream.setEncoding(encoding);
  }

  pause(): void {
    return this.stream.pause();
  }

  resume(): void {
    return this.stream.resume();
  }

  isPaused(): boolean {
    return this.stream.isPaused();
  }

  pipe(
    destination: Parameters<InstanceType<typeof StreamBuf>['pipe']>[0],
    _options?: unknown,
  ): void {
    // StreamBuf.pipe ignores options; signature kept for Readable parity
    return this.stream.pipe(destination);
  }

  unpipe(destination: Parameters<InstanceType<typeof StreamBuf>['unpipe']>[0]): void {
    return this.stream.unpipe(destination);
  }

  unshift(chunk?: unknown): never {
    return this.stream.unshift(chunk);
  }

  wrap(stream?: unknown): never {
    return this.stream.wrap(stream);
  }
}

export default {ZipWriter};
export {ZipWriter};
