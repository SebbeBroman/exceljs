import {EventEmitter} from './event-emitter.js';

export interface StutteredPipeOptions {
  bufSize?: number;
  autoPause?: boolean;
}

/** Minimal readable used by StutteredPipe. */
interface ReadableLike {
  on(event: string, listener: (...args: unknown[]) => void): unknown;
  read(size?: number): {length: number} | null | undefined;
}

/** Minimal writable used by StutteredPipe. */
interface WritableLike {
  write(data: unknown): unknown;
  end(): unknown;
}

// =============================================================================
// StutteredPipe - Used to slow down streaming so GC can get a look in
class StutteredPipe extends EventEmitter {
  readable: ReadableLike;
  writable: WritableLike;
  bufSize: number;
  autoPause: boolean;
  paused: boolean;
  eod: boolean;
  scheduled: ReturnType<typeof setImmediate> | null;

  constructor(
    readable: ReadableLike,
    writable: WritableLike,
    options?: StutteredPipeOptions,
  ) {
    super();

    options = options || {};

    this.readable = readable;
    this.writable = writable;
    this.bufSize = options.bufSize || 16384;
    this.autoPause = options.autoPause || false;

    this.paused = false;
    this.eod = false;
    this.scheduled = null;

    readable.on('end', () => {
      this.eod = true;
      writable.end();
    });

    // need to have some way to communicate speed of stream
    // back from the consumer
    readable.on('readable', () => {
      if (!this.paused) {
        this.resume();
      }
    });
    this._schedule();
  }

  pause(): void {
    this.paused = true;
  }

  resume(): void {
    if (!this.eod) {
      if (this.scheduled !== null) {
        clearImmediate(this.scheduled);
      }
      this._schedule();
    }
  }

  _schedule(): void {
    this.scheduled = setImmediate(() => {
      this.scheduled = null;
      if (!this.eod && !this.paused) {
        const data = this.readable.read(this.bufSize);
        if (data && data.length) {
          this.writable.write(data);

          if (!this.paused && !this.autoPause) {
            this._schedule();
          }
        } else if (!this.paused) {
          this._schedule();
        }
      }
    });
  }
}

export default StutteredPipe;
export {StutteredPipe};
