/**
 * Streaming ZIP reader built on fflate Unzip / UnzipInflate.
 * Yields entries as async-iterable objects (no unzipper / readable-stream).
 */
import {Buffer} from 'buffer';
import {Unzip, UnzipInflate} from 'fflate';
import {fromReadable} from './async-iterator.js';

/** Max decompressed chunks buffered per entry before pausing the input pump. */
const HIGH_WATER_MARK = 16;

/**
 * Copy into a detached Uint8Array. fflate retains subarray views of pushed chunks
 * (partial headers / file data), so sharing the Node Buffer pool is unsafe.
 * @param {unknown} chunk
 * @returns {Uint8Array}
 */
function toUint8Array(chunk) {
  if (chunk == null) {
    return new Uint8Array(0);
  }
  if (typeof chunk === 'string') {
    return new Uint8Array(Buffer.from(chunk));
  }
  if (chunk instanceof ArrayBuffer) {
    return new Uint8Array(chunk.slice(0));
  }
  if (ArrayBuffer.isView(chunk)) {
    // TypedArray / Buffer: constructor copies when given another TypedArray
    return new Uint8Array(chunk);
  }
  throw new Error('zip stream chunk must be Buffer, Uint8Array, ArrayBuffer, or string');
}

/**
 * Copy to a Buffer so later fflate reuse of internal arrays cannot corrupt data.
 * @param {Uint8Array} u8
 * @returns {Buffer}
 */
function u8ToBuffer(u8) {
  return Buffer.from(u8); // always copy
}

/**
 * One ZIP member. Read via async iteration or pipeTo(); call autodrain() to skip.
 */
export class ZipStreamEntry {
  /**
   * @param {import('fflate').UnzipFile} file
   * @param {{onPause?: () => void, onResume?: () => void}} [hooks]
   */
  constructor(file, hooks = {}) {
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

  _signal() {
    const wait = this._wait;
    this._wait = null;
    if (wait) wait();
  }

  _pauseIfNeeded() {
    if (this._queue.length >= HIGH_WATER_MARK && this._hooks.onPause) {
      this._hooks.onPause();
    }
  }

  _resumeIfNeeded() {
    if (this._queue.length < HIGH_WATER_MARK / 2 && this._hooks.onResume) {
      this._hooks.onResume();
    }
  }

  _finish(err) {
    if (err) this._error = err;
    if (!this._done) {
      this._done = true;
      this._resolveConsumed();
      if (this._hooks.onResume) this._hooks.onResume();
    }
    this._signal();
  }

  /**
   * Begin decompressing. Safe to call more than once.
   */
  start() {
    if (this._started) return;
    this._started = true;
    this._file.ondata = (err, dat, final) => {
      if (err) {
        this._finish(err);
        return;
      }
      if (dat && dat.length) {
        this._queue.push(u8ToBuffer(dat));
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
      this._finish(err);
    }
  }

  /**
   * Skip / drain without retaining payload (unzipper entry.autodrain() equivalent).
   * Starts inflate and discards so fflate does not retain compressed buffers forever.
   * @returns {Promise<void>}
   */
  autodrain() {
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
        this._finish(err);
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
   * @returns {AsyncGenerator<Buffer, void, unknown>}
   */
  async *[Symbol.asyncIterator]() {
    this.start();
    while (!this._done || this._queue.length > 0) {
      if (this._queue.length === 0) {
        // eslint-disable-next-line no-await-in-loop
        await new Promise(resolve => {
          this._wait = resolve;
        });
        if (this._error) throw this._error;
        continue;
      }
      const chunk = this._queue.shift();
      this._resumeIfNeeded();
      yield chunk;
    }
    if (this._error) throw this._error;
  }

  /**
   * Write all decompressed bytes to a Node writable stream and end it.
   * @param {import('stream').Writable} writable
   * @returns {Promise<void>}
   */
  async pipeTo(writable) {
    for await (const chunk of this) {
      if (writable.destroyed || writable.writableEnded) {
        throw new Error(`writable closed while piping zip entry: ${this.path}`);
      }
      if (!writable.write(chunk)) {
        // eslint-disable-next-line no-await-in-loop
        await new Promise((resolve, reject) => {
          const onDrain = () => {
            writable.removeListener('error', onError);
            resolve();
          };
          const onError = err => {
            writable.removeListener('drain', onDrain);
            reject(err);
          };
          writable.once('drain', onDrain);
          writable.once('error', onError);
        });
      }
    }
    await new Promise((resolve, reject) => {
      const onFinish = () => {
        writable.removeListener('error', onError);
        resolve();
      };
      const onError = err => {
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
 *
 * @param {AsyncIterable|import('stream').Readable} source
 * @returns {AsyncGenerator<ZipStreamEntry, void, unknown>}
 */
export async function* streamZipEntries(source) {
  const entryQueue = [];
  let notify = () => {};
  let inputError = null;
  let inputDone = false;
  let paused = false;
  let resumePump = () => {};

  const pause = () => {
    paused = true;
  };
  const resume = () => {
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
  unzipper.register(UnzipInflate);

  const pump = (async () => {
    try {
      for await (const chunk of fromReadable(source)) {
        if (inputError) return;
        while (paused) {
          // eslint-disable-next-line no-await-in-loop
          await new Promise(resolve => {
            resumePump = resolve;
          });
        }
        unzipper.push(toUint8Array(chunk), false);
      }
      while (paused) {
        // eslint-disable-next-line no-await-in-loop
        await new Promise(resolve => {
          resumePump = resolve;
        });
      }
      unzipper.push(new Uint8Array(0), true);
    } catch (err) {
      inputError = err;
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
        // eslint-disable-next-line no-await-in-loop
        await new Promise(resolve => {
          notify = resolve;
        });
        continue;
      }
      const entry = entryQueue.shift();
      yield entry;
      // Consumer finished this turn: drain if unread so inflate can complete.
      // eslint-disable-next-line no-await-in-loop
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
