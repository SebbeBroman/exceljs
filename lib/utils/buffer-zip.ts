/**
 * Collecting zip writer for buffered output (`XLSX.writeBuffer`).
 *
 * The buffered write path fully materializes every part (XML strings,
 * media buffers) before zipping, so fflate's streaming `Zip` only adds
 * event-loop hops: profiled ~50% idle on dense writes. Collecting parts
 * and running a single synchronous `zipSync` removes that overhead with
 * identical entry bytes. The streaming `StreamZipWriter` remains for the
 * `WorkbookWriter` / `writeFile` streaming APIs.
 */

import {strToU8, zipSync} from 'fflate';
import {asUint8Array, fromBase64, isBytes, toPublic} from './bytes.js';

export interface BufferZipAppendOptions {
  name: string;
  base64?: boolean;
}

function toEntryBytes(data: unknown, base64?: boolean): Uint8Array {
  if (base64 && typeof data === 'string') {
    return fromBase64(data);
  }
  if (typeof data === 'string') {
    return strToU8(data);
  }
  if (data instanceof Uint8Array) {
    return data;
  }
  if (isBytes(data)) {
    return asUint8Array(data as ArrayBuffer | ArrayBufferView);
  }
  return strToU8(String(data));
}

/** Minimal deflate-level resolution (mirrors StreamZipWriter.resolveLevel). */
export function resolveZipLevel(zipOptions?: unknown): number {
  const options = (zipOptions || {}) as {
    level?: number;
    zlib?: {level?: number};
    compression?: string;
    store?: boolean;
  };
  if (options.level != null) return options.level;
  if (options.zlib && options.zlib.level != null) return options.zlib.level;
  if (options.compression === 'STORE') return 0;
  if (options.store === true) return 0;
  // Default level 1: much faster than 6 with similar size order-of-magnitude
  return 1;
}

class BufferZipWriter {
  files: Record<string, Uint8Array> = Object.create(null);

  append(data: unknown, options: BufferZipAppendOptions): this {
    if (!options || !options.name) {
      throw new Error('zip.append requires options.name');
    }
    this.files[options.name] = toEntryBytes(data, options.base64);
    return this;
  }

  /** Compress all collected parts (level from `resolveZipLevel`). */
  toBytes(level: number): Uint8Array {
    const out = zipSync(this.files, {level: level as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9});
    return toPublic(out) as Uint8Array;
  }
}

export default BufferZipWriter;
export {BufferZipWriter};
