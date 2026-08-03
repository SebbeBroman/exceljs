/**
 * Tiny helpers so XML parse paths don't need readable-stream / PassThrough.
 */

/** Async-iterate string content in chunks (or one shot if small). */
export async function* stringChunks(content, chunkSize = 16 * 1024) {
  if (content == null) return;
  const str = typeof content === 'string' ? content : String(content);
  if (str.length <= chunkSize) {
    yield str;
    return;
  }
  for (let i = 0; i < str.length; i += chunkSize) {
    yield str.substring(i, i + chunkSize);
  }
}

/** Yield a single value (Buffer / Uint8Array / string). */
export async function* once(value) {
  yield value;
}

/**
 * Turn a Node-style readable (or any event emitter with data/end/error)
 * into an async iterable without PassThrough / readable-stream.
 */
export async function* fromReadable(stream) {
  if (stream == null) return;

  // Already async-iterable (Node Readable, our helpers, arrays, etc.)
  if (typeof stream[Symbol.asyncIterator] === 'function') {
    yield* stream;
    return;
  }
  if (typeof stream[Symbol.iterator] === 'function') {
    yield* stream;
    return;
  }

  // Legacy stream: .on('data') / 'end' / 'error'
  if (typeof stream.on === 'function') {
    const queue = [];
    let done = false;
    let error = null;
    let notify = () => {};

    const wake = () => {
      const n = notify;
      notify = () => {};
      n();
    };

    stream.on('data', chunk => {
      queue.push(chunk);
      wake();
    });
    stream.on('end', () => {
      done = true;
      wake();
    });
    stream.on('error', err => {
      error = err;
      done = true;
      wake();
    });

    if (typeof stream.resume === 'function') {
      stream.resume();
    }

    while (!done || queue.length > 0) {
      if (queue.length === 0) {
        // eslint-disable-next-line no-await-in-loop
        await new Promise(resolve => {
          notify = resolve;
        });
      }
      while (queue.length > 0) {
        yield queue.shift();
      }
      if (error) throw error;
    }
    return;
  }

  throw new Error('Cannot iterate value: expected async iterable or readable stream');
}
