import {EventEmitter} from './event-emitter.js';

// =============================================================================
// AutoDrain - kind of /dev/null
class AutoDrain extends EventEmitter {
  write(chunk: unknown): void {
    this.emit('data', chunk);
  }

  end(): void {
    this.emit('end');
  }
}

export default AutoDrain;
export {AutoDrain};
