/**
 * Streaming ZIP reader. Zip structure is fflate; DEFLATE payloads use Node zlib
 * (this entry is Node-only). Yields entries as async-iterable objects.
 */
import {createInflateRaw, type InflateRaw} from 'node:zlib';
import {Unzip} from 'fflate';
import type {UnzipFile} from 'fflate';
import {fromReadable} from './async-iterator.js';
import {from as bytesFrom, toPublic} from './bytes.js';
import type {Writable} from 'node:stream';

/** Buffers produced by zlib that the caller may retain without copying. */
const ownedZipChunks = new WeakSet<object>();

/**
 * Streaming raw DEFLATE (ZIP method 8) via Node's native inflater.
 * Registered with fflate Unzip in place of UnzipInflate.
 */
class NodeRawInflate {
  static compression = 8;
  ondata: (err: Error | null, data: Uint8Array | null, final: boolean) => void = () => {};
  private infl: InflateRaw;
  private failed = false;

  constructor() {
    this.infl = createInflateRaw();
    this.infl.on('data', (dat: Buffer) => {
      ownedZipChunks.add(dat.buffer);
      this.ondata(null, dat, false);
    });
    this.infl.on('end', () => {
      this.ondata(null, null, true);
    });
    this.infl.on('error', (err: Error) => {
      if (this.failed) return;
      this.failed = true;
      this.ondata(err, null, true);
    });
  }

  push(chunk: Uint8Array, final: boolean): void {
    if (this.failed) return;
    try {
      if (chunk && chunk.length) {
        // fflate hands out views of buffers it may reuse after push returns.
        const copy = Buffer.from(chunk.buffer, chunk.byteOffset, chunk.byteLength);
        this.infl.write(copy);
      }
      if (final) {
        this.infl.end();
      }
    } catch (err) {
      this.failed = true;
      this.ondata(err instanceof Error ? err : new Error(String(err)), null, final);
    }
  }
}

/** Max decompressed chunks buffered per entry before pausing the input pump. */
const HIGH_WATER_MARK = 16;

/**
 * Copy into a detached Uint8Array. fflate retains subarray views of pushed chunks
 * (partial headers / file data), so sharing the Node Buffer pool is unsafe.
 */
function toUint8Array(chunk: unknown): Uint8Array {
  if (chunk == null) {
    return new Uint8Array(0);
  }
  if (typeof chunk === 'string') {
    return bytesFrom(chunk, 'utf8');
  }
  if (chunk instanceof ArrayBuffer) {
    return new Uint8Array(chunk.slice(0));
  }
  if (ArrayBuffer.isView(chunk)) {
    // TypedArray / Buffer: constructor copies when given another TypedArray
    const view = chunk as ArrayBufferView;
    return new Uint8Array(view.buffer, view.byteOffset, view.byteLength).slice();
  }
  throw new Error('zip stream chunk must be Buffer, Uint8Array, ArrayBuffer, or string');
}

/**
 * Copy so later fflate reuse of internal arrays cannot corrupt data.
 */
function u8ToBuffer(u8: Uint8Array): Uint8Array {
  // always copy, then expose as public Buffer on Node
  return toPublic(u8.slice()) as Uint8Array;
}

export interface ZipStreamEntryHooks {
  onPause?: () => void;
  onResume?: () => void;
}

/**
 * One ZIP member. Read via async iteration or pipeTo(); call autodrain() to skip.
 */
export class ZipStreamEntry {
  path: string;
  name: string;
  _file: UnzipFile;
  _hooks: ZipStreamEntryHooks;
  _started: boolean;
  _queue: Uint8Array[];
  _done: boolean;
  _error: Error | null;
  _wait: (() => void) | null;
  _resolveConsumed: (() => void) | null;
  _consumed: Promise<void>;

  constructor(file: UnzipFile, hooks: ZipStreamEntryHooks = {}) {
    this.path = file.name;
    this.name = file.name;
    this._file = file;
    this._hooks = hooks;
    this._started = false;
    this._queue = [];
    this._done = false;
    this._error = null;
    this._wait = null;
    this._resolveConsumed = null;
    this._consumed = new Promise(resolve => {
      this._resolveConsumed = resolve;
    });
  }

  _signal(): void {
    const wait = this._wait;
    this._wait = null;
    if (wait) wait();
  }

  _pauseIfNeeded(): void {
    if (this._queue.length >= HIGH_WATER_MARK && this._hooks.onPause) {
      this._hooks.onPause();
    }
  }

  _resumeIfNeeded(): void {
    if (this._queue.length < HIGH_WATER_MARK / 2 && this._hooks.onResume) {
      this._hooks.onResume();
    }
  }

  _finish(err: Error | null): void {
    if (err) this._error = err;
    if (!this._done) {
      this._done = true;
      this._resolveConsumed!();
      if (this._hooks.onResume) this._hooks.onResume();
    }
    this._signal();
  }

  /**
   * Begin decompressing. Safe to call more than once.
   */
  start(): void {
    if (this._started) return;
    this._started = true;
    this._file.ondata = (err, dat, final) => {
      if (err) {
        this._finish(err);
        return;
      }
      if (dat && dat.length) {
        this._queue.push(ownedZipChunks.has(dat.buffer) ? dat : u8ToBuffer(dat));
        this._pauseIfNeeded();
      }
      if (final) {
        this._finish(null);
      } else {
        this._signal();
      }
    };
    try {
      this._file.start();
    } catch (err) {
      this._finish(err as Error);
    }
  }

  /**
   * Skip / drain without retaining payload (unzipper entry.autodrain() equivalent).
   * Starts inflate and discards so fflate does not retain compressed buffers forever.
   */
  autodrain(): Promise<void> {
    if (this._done) {
      return this._consumed;
    }
    if (!this._started) {
      this._started = true;
      this._file.ondata = (err, _dat, final) => {
        if (err || final) this._finish(err || null);
      };
      try {
        this._file.start();
      } catch (err) {
        this._finish(err as Error);
      }
    } else {
      // Already streaming: drop buffered data and discard further chunks.
      this._queue.length = 0;
      this._resumeIfNeeded();
      this._file.ondata = (err, _dat, final) => {
        if (err || final) this._finish(err || null);
      };
    }
    return this._consumed;
  }

  /**
   * Async-iterate decompressed bytes as Buffers.
   */
  async *[Symbol.asyncIterator](): AsyncGenerator<Uint8Array, void, unknown> {
    this.start();
    while (!this._done || this._queue.length > 0) {
      if (this._queue.length === 0) {
        await new Promise<void>(resolve => {
          this._wait = resolve;
        });
        if (this._error) throw this._error;
        continue;
      }
      const chunk = this._queue.shift()!;
      this._resumeIfNeeded();
      yield chunk;
    }
    if (this._error) throw this._error;
  }

  /**
   * Write all decompressed bytes to a Node writable stream and end it.
   */
  async pipeTo(writable: Writable): Promise<void> {
    for await (const chunk of this) {
      if (writable.destroyed || writable.writableEnded) {
        throw new Error(`writable closed while piping zip entry: ${this.path}`);
      }
      if (!writable.write(chunk)) {
        await new Promise<void>((resolve, reject) => {
          const onDrain = (): void => {
            writable.removeListener('error', onError);
            resolve();
          };
          const onError = (err: Error): void => {
            writable.removeListener('drain', onDrain);
            reject(err);
          };
          writable.once('drain', onDrain);
          writable.once('error', onError);
        });
      }
    }
    await new Promise<void>((resolve, reject) => {
      const onFinish = (): void => {
        writable.removeListener('error', onError);
        resolve();
      };
      const onError = (err: Error): void => {
        writable.removeListener('finish', onFinish);
        reject(err);
      };
      writable.once('finish', onFinish);
      writable.once('error', onError);
      writable.end();
    });
  }
}

/**
 * Turn a Node readable (or async iterable) of zip bytes into an async generator of entries.
 * After each yield, if the consumer did not fully read the entry, remaining data is drained.
 */
export async function* streamZipEntries(
  source: unknown,
): AsyncGenerator<ZipStreamEntry, void, unknown> {
  const entryQueue: ZipStreamEntry[] = [];
  let notify: () => void = () => {};
  let inputError: Error | null = null;
  let inputDone = false;
  let paused = false;
  let resumePump: () => void = () => {};

  const pause = (): void => {
    paused = true;
  };
  const resume = (): void => {
    if (paused) {
      paused = false;
      const r = resumePump;
      resumePump = () => {};
      r();
    }
  };

  const unzipper = new Unzip(file => {
    const entry = new ZipStreamEntry(file, {onPause: pause, onResume: resume});
    entryQueue.push(entry);
    notify();
  });
  unzipper.register(NodeRawInflate);

  const pump = (async () => {
    try {
      for await (const chunk of fromReadable(source)) {
        if (inputError) return;
        while (paused) {
          await new Promise<void>(resolve => {
            resumePump = resolve;
          });
        }
        unzipper.push(toUint8Array(chunk), false);
      }
      while (paused) {
        await new Promise<void>(resolve => {
          resumePump = resolve;
        });
      }
      unzipper.push(new Uint8Array(0), true);
    } catch (err) {
      inputError = err as Error;
    } finally {
      inputDone = true;
      notify();
      resume();
    }
  })();

  try {
    while (!inputDone || entryQueue.length > 0) {
      if (entryQueue.length === 0) {
        if (inputError) throw inputError;

        await new Promise<void>(resolve => {
          notify = resolve;
        });
        continue;
      }
      const entry = entryQueue.shift()!;
      yield entry;
      // Consumer finished this turn: drain if unread so inflate can complete.

      await entry.autodrain();
      if (entry._error) throw entry._error;
    }
    if (inputError) throw inputError;
    await pump;
  } catch (err) {
    for (const e of entryQueue) {
      e.autodrain();
    }
    resume();
    throw err;
  }
}

export default streamZipEntries;
