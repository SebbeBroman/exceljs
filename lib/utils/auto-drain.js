import {EventEmitter} from './event-emitter.js';

// =============================================================================
// AutoDrain - kind of /dev/null
class AutoDrain extends EventEmitter {
  write(chunk) {
    this.emit('data', chunk);
  }

  end() {
    this.emit('end');
  }
}

export default AutoDrain;
export {AutoDrain};
