import {EventEmitter} from 'events';

// =============================================================================
// StreamBase64 - A utility to convert to/from base64 stream
// Note: does not buffer data, must be piped
// (No readable-stream — EventEmitter only, process-free.)
class StreamBase64 extends EventEmitter {
  constructor() {
    super();

    // consuming pipe streams go here
    this.pipes = [];
  }

  // writable
  write(/* data, encoding */) {
    return true;
  }

  cork() {}

  uncork() {}

  end(/* chunk, encoding, callback */) {}

  // readable
  read(/* size */) {}

  setEncoding(encoding) {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  }

  pause() {}

  resume() {}

  isPaused() {}

  pipe(destination) {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
  }

  unpipe(destination) {
    // remove destination from pipe list
    this.pipes = this.pipes.filter(pipe => pipe !== destination);
  }

  unshift(/* chunk */) {
    throw new Error('Not Implemented');
  }

  wrap(/* stream */) {
    throw new Error('Not Implemented');
  }
}

export default StreamBase64;
export {StreamBase64};
