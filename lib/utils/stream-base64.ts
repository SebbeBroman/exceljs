import {EventEmitter} from './event-emitter.js';

/** Minimal writable pipe destination. */
interface PipeDestination {
  write?(chunk: unknown, encoding?: string, callback?: () => void): unknown;
  end?(chunk?: unknown, encoding?: string, callback?: () => void): unknown;
}

// =============================================================================
// StreamBase64 - A utility to convert to/from base64 stream
// Note: does not buffer data, must be piped
// (No readable-stream — EventEmitter only, process-free.)
class StreamBase64 extends EventEmitter {
  pipes: PipeDestination[];
  encoding: string | undefined;

  constructor() {
    super();

    // consuming pipe streams go here
    this.pipes = [];
  }

  // writable
  write(/* data, encoding */): boolean {
    return true;
  }

  cork(): void {}

  uncork(): void {}

  end(/* chunk, encoding, callback */): void {}

  // readable
  read(/* size */): void {}

  setEncoding(encoding: string): void {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  }

  pause(): void {}

  resume(): void {}

  isPaused(): void {}

  pipe(destination: PipeDestination): void {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
  }

  unpipe(destination: PipeDestination): void {
    // remove destination from pipe list
    this.pipes = this.pipes.filter(pipe => pipe !== destination);
  }

  unshift(/* chunk */): never {
    throw new Error('Not Implemented');
  }

  wrap(/* stream */): never {
    throw new Error('Not Implemented');
  }
}

export default StreamBase64;
export {StreamBase64};
