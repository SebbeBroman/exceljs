import {createRequire} from 'node:module';
import {unzipSync} from 'fflate';
import {asUint8Array, toPublic, toString as bytesToString} from './bytes.js';
import {isNode} from './env.js';

/**
 * @param data
 * @returns
 */
function toUint8Array(data: ArrayBuffer | ArrayBufferView): Uint8Array {
  return asUint8Array(data);
}

/** Optional Node zlib.inflateRawSync — much faster than pure-JS inflate. */
type InflateRawSync = (data: Uint8Array, opts?: {maxOutputLength?: number}) => Uint8Array;
let inflateRawSync: InflateRawSync | null | undefined;

function getInflateRawSync(): InflateRawSync | null {
  if (inflateRawSync !== undefined) return inflateRawSync;
  if (!isNode()) {
    inflateRawSync = null;
    return null;
  }
  try {
    // createRequire is browser-shimmed to throw; only called when isNode().
    const req = createRequire(import.meta.url);
    const zlib = req('node:zlib') as {inflateRawSync: InflateRawSync};
    inflateRawSync = zlib.inflateRawSync.bind(zlib);
  } catch {
    inflateRawSync = null;
  }
  return inflateRawSync;
}

/**
 * Parse a classic ZIP (stored / deflate) with Node's native inflate.
 * Throws on ZIP64 / unsupported methods so callers can fall back to fflate.
 *
 * Security policy: CRC32 is not re-verified on the native path (matches
 * fflate `unzipSync` behavior for corrupt-entry tolerance); callers should
 * treat a successful parse as structurally valid, not integrity-proof.
 * Decompression-bomb guard: total *actual* uncompressed output is capped at
 * MAX_TOTAL_UNCOMPRESSED (512 MiB) and entry count at MAX_ENTRIES.
 * Header `uncompSize` is never trusted: inflate is bounded by the remaining
 * budget and the real output length is accounted afterwards.
 */
export const ZIP_LIMITS = {
  maxEntries: 10_000,
  maxTotalUncompressed: 512 * 1024 * 1024,
} as const;
const MAX_TOTAL_UNCOMPRESSED = ZIP_LIMITS.maxTotalUncompressed;
const MAX_ENTRIES = ZIP_LIMITS.maxEntries;

/** Marker for decompression-guard rejections — must never fall through to fflate. */
export const ZIP_LIMIT_EXCEEDED = 'ZIP_LIMIT_EXCEEDED';

function zipLimitError(message: string): Error {
  const err = new Error(message);
  (err as {code?: string}).code = ZIP_LIMIT_EXCEEDED;
  return err;
}

export function isZipLimitError(err: unknown): boolean {
  return (
    err instanceof Error &&
    ((err as {code?: string}).code === ZIP_LIMIT_EXCEEDED || /exceeds limit/i.test(err.message))
  );
}

/** Cheap EOCD entry-count pre-check so the fflate path can fail fast without inflating. */
function zipEntryCountHint(data: Uint8Array): number | null {
  try {
    if (data.length < 22) return null;
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    let eocd = data.length - 22;
    const minEocd = Math.max(0, data.length - 22 - 0xffff);
    while (eocd >= minEocd) {
      if (view.getUint32(eocd, true) === 0x06054b50) break;
      eocd--;
    }
    if (eocd < minEocd) return null;
    const totalEntries = view.getUint16(eocd + 10, true);
    if (totalEntries === 0xffff) return null; // ZIP64 — unknown here
    return totalEntries;
  } catch {
    return null;
  }
}

/** Enforce entry-count + total-size caps on an already-inflated file map (fflate path). */
function enforceUnzipLimits(files: Record<string, Uint8Array>): void {
  const names = Object.keys(files);
  if (names.length > MAX_ENTRIES) {
    throw zipLimitError(`ZIP entry count ${names.length} exceeds limit ${MAX_ENTRIES}`);
  }
  let total = 0;
  for (const name of names) {
    total += files[name]!.byteLength;
    if (total > MAX_TOTAL_UNCOMPRESSED) {
      throw zipLimitError('ZIP uncompressed size exceeds 512 MiB limit');
    }
  }
}
function unzipNative(data: Uint8Array): Record<string, Uint8Array> {
  const inflate = getInflateRawSync();
  if (!inflate) {
    throw new Error('native inflate unavailable');
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  // EOCD is at the end (optionally preceded by a comment). Scan backwards.
  let eocd = data.length - 22;
  const minEocd = Math.max(0, data.length - 22 - 0xffff);
  while (eocd >= minEocd) {
    if (view.getUint32(eocd, true) === 0x06054b50) break;
    eocd--;
  }
  if (eocd < minEocd) {
    throw new Error('ZIP EOCD not found');
  }

  const totalEntries = view.getUint16(eocd + 10, true);
  let cdOffset = view.getUint32(eocd + 16, true);
  if (totalEntries > MAX_ENTRIES) {
    throw zipLimitError(`ZIP entry count ${totalEntries} exceeds limit ${MAX_ENTRIES}`);
  }
  // ZIP64: offsets of 0xffffffff need the ZIP64 EOCD locator — let fflate handle those.
  if (cdOffset === 0xffffffff || totalEntries === 0xffff) {
    throw new Error('ZIP64 not handled by native path');
  }

  const files: Record<string, Uint8Array> = Object.create(null);
  let totalUncompressed = 0;
  // File names in OOXML are ASCII/UTF-8; decode once per entry.
  const nameDecoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : null;

  for (let i = 0; i < totalEntries; i++) {
    if (view.getUint32(cdOffset, true) !== 0x02014b50) {
      throw new Error('bad ZIP central directory signature');
    }
    const method = view.getUint16(cdOffset + 10, true);
    const compSize = view.getUint32(cdOffset + 20, true);
    const uncompSize = view.getUint32(cdOffset + 24, true);
    const nameLen = view.getUint16(cdOffset + 28, true);
    const extraLen = view.getUint16(cdOffset + 30, true);
    const commentLen = view.getUint16(cdOffset + 32, true);
    const localOffset = view.getUint32(cdOffset + 42, true);
    if (compSize === 0xffffffff || uncompSize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error('ZIP64 entry not handled by native path');
    }

    const nameBytes = data.subarray(cdOffset + 46, cdOffset + 46 + nameLen);
    let name: string;
    if (nameDecoder) {
      name = nameDecoder.decode(nameBytes);
    } else {
      name = '';
      for (let j = 0; j < nameBytes.length; j++) name += String.fromCharCode(nameBytes[j]);
    }
    cdOffset += 46 + nameLen + extraLen + commentLen;

    if (!name || name.endsWith('/')) continue;

    const lh = localOffset;
    if (view.getUint32(lh, true) !== 0x04034b50) {
      throw new Error(`bad ZIP local header for ${name}`);
    }
    const lhNameLen = view.getUint16(lh + 26, true);
    const lhExtraLen = view.getUint16(lh + 28, true);
    const dataStart = lh + 30 + lhNameLen + lhExtraLen;
    const comp = data.subarray(dataStart, dataStart + compSize);

    if (method === 0) {
      // Stored — copy so callers cannot mutate the source archive buffer.
      // Account actual bytes, not header claims.
      if (totalUncompressed + compSize > MAX_TOTAL_UNCOMPRESSED) {
        throw zipLimitError('ZIP uncompressed size exceeds 512 MiB limit');
      }
      totalUncompressed += compSize;
      files[name] = comp.slice();
    } else if (method === 8) {
      // Never trust the header size: bound inflate by the remaining budget and
      // account the real output length afterwards (catches lying headers).
      const remaining = MAX_TOTAL_UNCOMPRESSED - totalUncompressed;
      if (remaining <= 0) {
        throw zipLimitError('ZIP uncompressed size exceeds 512 MiB limit');
      }
      let out: Uint8Array;
      try {
        out = inflate(comp, {maxOutputLength: remaining});
      } catch (err) {
        if (isZipLimitError(err)) throw err;
        // Node throws when output exceeds maxOutputLength — normalize to a
        // guard error so callers never fall through to the uncapped path.
        const msg = err instanceof Error ? err.message : String(err);
        if (/maxOutputLength|too large|output length|ERR_BUFFER/i.test(msg)) {
          throw zipLimitError('ZIP uncompressed size exceeds 512 MiB limit');
        }
        throw err;
      }
      const bytes = out instanceof Uint8Array ? out : new Uint8Array(out as ArrayBuffer);
      totalUncompressed += bytes.byteLength;
      if (totalUncompressed > MAX_TOTAL_UNCOMPRESSED) {
        throw zipLimitError('ZIP uncompressed size exceeds 512 MiB limit');
      }
      // Node returns Buffer (Uint8Array subclass); keep as-is (no extra copy).
      files[name] = bytes;
    } else {
      throw new Error(`unsupported ZIP method ${method} for ${name}`);
    }
  }

  return files;
}

/**
 * Unzip an OOXML package into path → bytes.
 * Prefer Node zlib (native) then fflate `unzipSync`. Avoid fflate's async `unzip`:
 * it yields between inflate chunks and is ~10–15× slower for typical xlsx sizes.
 *
 * @param data
 * @returns
 */
export function unzipToFiles(
  data: ArrayBuffer | ArrayBufferView,
): Promise<Record<string, Uint8Array>> {
  const u8 = toUint8Array(data);
  try {
    if (getInflateRawSync()) {
      try {
        return Promise.resolve(unzipNative(u8));
      } catch (err) {
        // Guard rejections must propagate — falling through to uncapped fflate
        // would defeat the bomb protection.
        if (isZipLimitError(err)) throw err;
        // ZIP64 / exotic — fall through to fflate
      }
    }
    // fflate path (browser + ZIP64 fallback): fail fast on entry count, then
    // enforce total-size caps on the inflated output before returning it.
    const hint = zipEntryCountHint(u8);
    if (hint != null && hint > MAX_ENTRIES) {
      throw zipLimitError(`ZIP entry count ${hint} exceeds limit ${MAX_ENTRIES}`);
    }
    const files = unzipSync(u8) as Record<string, Uint8Array>;
    enforceUnzipLimits(files);
    return Promise.resolve(files);
  } catch (err) {
    return Promise.reject(err);
  }
}

/**
 * Decode zip entry bytes as UTF-8 text (XML parts).
 * Uses the shared TextDecoder in bytes.ts (module-level, no per-call alloc).
 * @param u8
 */
export function entryToString(u8: Uint8Array): string {
  return bytesToString(u8);
}

/**
 * @param u8
 * @returns public buffer (Node Buffer when available) — zero-copy view when possible
 */
export function entryToBuffer(u8: Uint8Array | ArrayBufferView): Uint8Array {
  return toPublic(u8 instanceof Uint8Array ? u8 : asUint8Array(u8)) as Uint8Array;
}
