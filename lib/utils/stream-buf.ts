import {EventEmitter} from './event-emitter.js';
import type {EventEmitter as EventEmitterInstance} from './event-emitter.js';
import utils from './utils.js';
import StringBuf from './string-buf.js';
import {nextTick} from './env.js';
import {alloc, asUint8Array, concat, copy, from, isBytes, toPublic, utf8Encode} from './bytes.js';

export interface StreamBufOptions {
  bufSize?: number;
  batch?: boolean;
}

/** Chunk that can be copied into a ReadWriteBuf. */
interface DataChunk {
  length: number;
  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number;
  toBuffer(): Uint8Array;
}

/** Writable pipe destination. */
interface PipeDestination {
  write(chunk: Uint8Array | Buffer, callback?: (error?: Error | null) => void): unknown;
  end(): unknown;
}

// =============================================================================
// data chunks - encapsulating incoming data
class StringChunk implements DataChunk {
  _data: string | String | ArrayBuffer;
  _encoding: string | undefined;
  _buffer: Uint8Array | undefined;

  constructor(data: string | String | ArrayBuffer, encoding?: string) {
    this._data = data;
    this._encoding = encoding;
  }

  get length(): number {
    return this.toBuffer().length;
  }

  // copy to target buffer
  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    return copy(this.toBuffer(), target, targetOffset, offset, offset + length);
  }

  toBuffer(): Uint8Array {
    if (!this._buffer) {
      if (typeof this._data === 'string' || this._data instanceof String) {
        this._buffer = from(String(this._data), this._encoding || 'utf8');
      } else if (this._data instanceof ArrayBuffer) {
        this._buffer = from(this._data);
      } else {
        this._buffer = utf8Encode(String(this._data));
      }
    }
    return this._buffer;
  }
}

class StringBufChunk implements DataChunk {
  _data: StringBuf;

  constructor(data: StringBuf) {
    this._data = data;
  }

  get length(): number {
    return this._data.length;
  }

  // copy to target buffer
  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    return copy(this._data._buf, target, targetOffset, offset, offset + length);
  }

  toBuffer(): Uint8Array {
    return this._data.toBuffer();
  }
}

class BufferChunk implements DataChunk {
  _data: Uint8Array;

  constructor(data: Uint8Array) {
    this._data = data;
  }

  get length(): number {
    return this._data.length;
  }

  // copy to target buffer
  copy(target: Uint8Array, targetOffset: number, offset: number, length: number): number {
    return copy(this._data, target, targetOffset, offset, offset + length);
  }

  toBuffer(): Uint8Array {
    return this._data;
  }
}

// =============================================================================
// ReadWriteBuf - a single buffer supporting simple read-write
class ReadWriteBuf {
  size: number;
  buffer: Uint8Array;
  iRead: number;
  iWrite: number;

  constructor(size: number) {
    this.size = size;
    // the buffer
    this.buffer = alloc(size);
    // read index
    this.iRead = 0;
    // write index
    this.iWrite = 0;
  }

  toBuffer(): Uint8Array {
    if (this.iRead === 0 && this.iWrite === this.size) {
      return this.buffer;
    }

    const buf = alloc(this.iWrite - this.iRead);
    copy(this.buffer, buf, 0, this.iRead, this.iWrite);
    return buf;
  }

  get length(): number {
    return this.iWrite - this.iRead;
  }

  get eod(): boolean {
    return this.iRead === this.iWrite;
  }

  get full(): boolean {
    return this.iWrite === this.size;
  }

  read(size?: number): Uint8Array | null {
    let buf: Uint8Array;
    // read size bytes from buffer and return buffer
    if (size === 0) {
      // special case - return null if no data requested
      return null;
    }

    if (size === undefined || size >= this.length) {
      // if no size specified or size is at least what we have then return all of the bytes
      buf = this.toBuffer();
      this.iRead = this.iWrite;
      return buf;
    }

    // otherwise return a chunk
    buf = alloc(size);
    copy(this.buffer, buf, 0, this.iRead, this.iRead + size);
    this.iRead += size;
    return buf;
  }

  write(chunk: DataChunk, offset: number, length: number): number {
    // write as many bytes from data from optional source offset
    // and return number of bytes written
    const size = Math.min(length, this.size - this.iWrite);
    chunk.copy(this.buffer, this.iWrite, offset, size);
    this.iWrite += size;
    return size;
  }
}

// =============================================================================
// StreamBuf - a multi-purpose read-write stream
//  As MemBuf - write as much data as you like. Then call toBuffer() to consolidate
//  As StreamHub - pipe to multiple writables
//  As readable stream - feed data into the writable part and have some other code read from it.

// Note: Not sure why but StreamBuf does not like JS "class" sugar. It fails the
// integration tests
export interface StreamBuf extends EventEmitterInstance {
  bufSize: number;
  buffers: ReadWriteBuf[];
  batch: boolean;
  corked: boolean;
  inPos: number;
  outPos: number;
  pipes: PipeDestination[];
  paused: boolean;
  encoding: string | null;

  toBuffer(): Uint8Array | null;
  _getWritableBuffer(): ReadWriteBuf;
  _pipe(chunk: DataChunk): Promise<void[]>;
  _writeToBuffers(chunk: DataChunk): void;
  write(
    data: unknown,
    encoding?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void,
  ): Promise<boolean> | boolean;
  cork(): void;
  _flush(): void;
  uncork(): void;
  end(
    chunk?: unknown,
    encoding?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void,
  ): void;
  read(size?: number): Uint8Array | null | undefined;
  setEncoding(encoding: string): void;
  pause(): void;
  resume(): void;
  isPaused(): boolean;
  pipe(destination: PipeDestination): void;
  unpipe(destination: PipeDestination): void;
  unshift(chunk?: unknown): never;
  wrap(stream?: unknown): never;
}

export interface StreamBufConstructor {
  new (options?: StreamBufOptions): StreamBuf;
  (options?: StreamBufOptions): StreamBuf;
  prototype: StreamBuf;
}

const StreamBuf = function StreamBuf(
  this: StreamBuf,
  options?: StreamBufOptions,
) {
  EventEmitter.call(this);
  options = options || {};
  this.bufSize = options.bufSize || 1024 * 1024;
  this.buffers = [];

  // batch mode fills a buffer completely before passing the data on
  // to pipes or 'readable' event listeners
  this.batch = options.batch || false;

  this.corked = false;
  // where in the current writable buffer we're up to
  this.inPos = 0;

  // where in the current readable buffer we've read up to
  this.outPos = 0;

  // consuming pipe streams go here
  this.pipes = [];

  // controls emit('data')
  this.paused = false;

  this.encoding = null;
} as StreamBufConstructor;

// EventEmitter only — not readable-stream (which depends on process.nextTick)
utils.inherits(StreamBuf, EventEmitter, {
  toBuffer(this: StreamBuf) {
    switch (this.buffers.length) {
      case 0:
        return null;
      case 1:
        return toPublic(this.buffers[0].toBuffer());
      default:
        return toPublic(concat(this.buffers.map(rwBuf => rwBuf.toBuffer())));
    }
  },

  // writable
  // event drain - if write returns false (which it won't), indicates when safe to write again.
  // finish - end() has been called
  // pipe(src) - pipe() has been called on readable
  // unpipe(src) - unpipe() has been called on readable
  // error - duh

  _getWritableBuffer(this: StreamBuf) {
    if (this.buffers.length) {
      const last = this.buffers[this.buffers.length - 1];
      if (!last.full) {
        return last;
      }
    }
    const buf = new ReadWriteBuf(this.bufSize);
    this.buffers.push(buf);
    return buf;
  },

  async _pipe(this: StreamBuf, chunk: DataChunk) {
    const write = function (pipe: PipeDestination) {
      return new Promise<void>(resolve => {
        pipe.write(chunk.toBuffer() as Buffer, () => {
          resolve();
        });
      });
    };
    await Promise.all(this.pipes.map(write));
  },
  _writeToBuffers(this: StreamBuf, chunk: DataChunk) {
    let inPos = 0;
    const inLen = chunk.length;
    while (inPos < inLen) {
      // find writable buffer
      const buffer = this._getWritableBuffer();

      // write some data
      inPos += buffer.write(chunk, inPos, inLen - inPos);
    }
  },
  async write(
    this: StreamBuf,
    data: unknown,
    encoding?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void,
  ) {
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || utils.nop;

    // encapsulate data into a chunk
    let chunk: DataChunk;
    if (data instanceof StringBuf) {
      chunk = new StringBufChunk(data);
    } else if (isBytes(data)) {
      // Uint8Array and Node Buffer
      chunk = new BufferChunk(data);
    } else if (ArrayBuffer.isView(data)) {
      // other typed arrays from fflate / browser APIs
      chunk = new BufferChunk(asUint8Array(data));
    } else if (typeof data === 'string' || data instanceof String || data instanceof ArrayBuffer) {
      chunk = new StringChunk(data, encoding as string | undefined);
    } else {
      throw new Error('Chunk must be one of type String, Buffer, Uint8Array or StringBuf.');
    }

    // now, do something with the chunk
    if (this.pipes.length) {
      if (this.batch) {
        this._writeToBuffers(chunk);
        while (!this.corked && this.buffers.length > 1) {
          this._pipe(this.buffers.shift() as unknown as DataChunk);
        }
      } else if (!this.corked) {
        await this._pipe(chunk);
        callback();
      } else {
        this._writeToBuffers(chunk);
        nextTick(callback);
      }
    } else {
      if (!this.paused) {
        this.emit('data', toPublic(chunk.toBuffer()));
      }

      this._writeToBuffers(chunk);
      this.emit('readable');
    }

    return true;
  },
  cork(this: StreamBuf) {
    this.corked = true;
  },
  _flush(this: StreamBuf /* destination */) {
    // if we have comsumers...
    if (this.pipes.length) {
      // and there's stuff not written
      while (this.buffers.length) {
        this._pipe(this.buffers.shift() as unknown as DataChunk);
      }
    }
  },
  uncork(this: StreamBuf) {
    this.corked = false;
    this._flush();
  },
  end(
    this: StreamBuf,
    chunk?: unknown,
    encoding?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void,
  ) {
    const writeComplete = (error?: Error | null) => {
      if (error) {
        if (callback) callback(error);
      } else {
        this._flush();
        this.pipes.forEach(pipe => {
          pipe.end();
        });
        this.emit('finish');
      }
    };
    if (chunk) {
      void this.write(chunk, encoding, writeComplete);
    } else {
      writeComplete();
    }
  },

  // readable
  // event readable - some data is now available
  // event data - switch to flowing mode - feeds chunks to handler
  // event end - no more data
  // event close - optional, indicates upstream close
  // event error - duh
  read(this: StreamBuf, size?: number) {
    let buffers: Uint8Array[];
    // read min(buffer, size || infinity)
    if (size) {
      buffers = [];
      while (size && this.buffers.length && !this.buffers[0].eod) {
        const first = this.buffers[0];
        const buffer = first.read(size)!;
        size -= buffer.length;
        buffers.push(buffer);
        if (first.eod && first.full) {
          this.buffers.shift();
        }
      }
      return toPublic(concat(buffers));
    }

    buffers = this.buffers.map(buf => buf.toBuffer()).filter(Boolean);
    this.buffers = [];
    return toPublic(concat(buffers));
  },
  setEncoding(this: StreamBuf, encoding: string) {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  },
  pause(this: StreamBuf) {
    this.paused = true;
  },
  resume(this: StreamBuf) {
    this.paused = false;
  },
  isPaused(this: StreamBuf) {
    return !!this.paused;
  },
  pipe(this: StreamBuf, destination: PipeDestination) {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
    if (!this.paused && this.buffers.length) {
      this.end();
    }
  },
  unpipe(this: StreamBuf, destination: PipeDestination) {
    // remove destination from pipe list
    this.pipes = this.pipes.filter(pipe => pipe !== destination);
  },
  unshift(/* chunk */) {
    // some numpty has read some data that's not for them and they want to put it back!
    // Might implement this some day
    throw new Error('Not Implemented');
  },
  wrap(/* stream */) {
    // not implemented
    throw new Error('Not Implemented');
  },
});

export default StreamBuf;
export {StreamBuf};
