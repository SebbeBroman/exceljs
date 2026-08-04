import {EventEmitter} from '../utils/event-emitter.js';

export interface LineBufferOptions {
  encoding?: string;
}

class LineBuffer extends EventEmitter {
  encoding: string | undefined;
  buffer: string | null;
  corked: boolean;
  queue: string[];

  constructor(options: LineBufferOptions = {}) {
    super();

    this.encoding = options.encoding;

    this.buffer = null;

    // part of cork/uncork
    this.corked = false;
    this.queue = [];
  }

  // Events:
  //  line: here is a line
  //  done: all lines emitted

  write(chunk: string): boolean {
    // find line or lines in chunk and emit them if not corked
    // or queue them if corked
    const data = this.buffer ? this.buffer + chunk : chunk;
    const lines = data.split(/\r?\n/g);

    // save the last line
    this.buffer = lines.pop() ?? null;

    // Preserve original forEach(function) this-binding (none / undefined in ESM)
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    lines.forEach(function (line: string) {
      // Original used unbound `this`; keep identical control flow via self.
      // Runtime: original was broken for corked in strict ESM; we preserve intended behavior.
      if (self.corked) {
        self.queue.push(line);
      } else {
        self.emit('line', line);
      }
    });

    return !this.corked;
  }

  cork(): void {
    this.corked = true;
  }

  uncork(): void {
    this.corked = false;
    this._flush();

    // tell the source I'm ready again
    this.emit('drain');
  }

  setDefaultEncoding(): void {
    // ?
  }

  end(): void {
    if (this.buffer) {
      this.emit('line', this.buffer);
      this.buffer = null;
    }
    this.emit('done');
  }

  _flush(): void {
    if (!this.corked) {
      this.queue.forEach(line => {
        this.emit('line', line);
      });
      this.queue = [];
    }
  }
}

export default LineBuffer;
export {LineBuffer};
