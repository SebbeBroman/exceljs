/**
 * Tiny helpers so XML parse paths don't need readable-stream / PassThrough.
 */

/**
 * Async-iterate string content for SAX.
 * When the full string is already in memory (post-unzip), yield once — chunking
 * only adds async-generator overhead for the buffered load path.
 * Pass `chunkSize` for true streaming scenarios that need bounded writes.
 */
export async function* stringChunks(
  content: unknown,
  chunkSize?: number,
): AsyncGenerator<string, void, unknown> {
  if (content == null) return;
  const str = typeof content === 'string' ? content : String(content);
  if (chunkSize == null || chunkSize <= 0 || str.length <= chunkSize) {
    yield str;
    return;
  }
  for (let i = 0; i < str.length; i += chunkSize) {
    yield str.substring(i, i + chunkSize);
  }
}

/** Yield a single value (Buffer / Uint8Array / string). */
export async function* once<T>(value: T): AsyncGenerator<T, void, unknown> {
  yield value;
}

/** Minimal readable / event-emitter surface. */
interface ReadableLike {
  on?(event: string, listener: (...args: unknown[]) => void): unknown;
  resume?(): unknown;
  [Symbol.asyncIterator]?: () => AsyncIterator<unknown>;
  [Symbol.iterator]?: () => Iterator<unknown>;
}

/**
 * Turn a Node-style readable (or any event emitter with data/end/error)
 * into an async iterable without PassThrough / readable-stream.
 */
export async function* fromReadable(stream: unknown): AsyncGenerator<unknown, void, unknown> {
  if (stream == null) return;

  const s = stream as ReadableLike;

  // Already async-iterable (Node Readable, our helpers, arrays, etc.)
  if (typeof s[Symbol.asyncIterator] === 'function') {
    yield* s as AsyncIterable<unknown>;
    return;
  }
  if (typeof s[Symbol.iterator] === 'function') {
    yield* s as Iterable<unknown>;
    return;
  }

  // Legacy stream: .on('data') / 'end' / 'error'
  if (typeof s.on === 'function') {
    const queue: unknown[] = [];
    let done = false;
    let error: unknown = null;
    let notify: () => void = () => {};

    const wake = (): void => {
      const n = notify;
      notify = () => {};
      n();
    };

    s.on('data', (chunk: unknown) => {
      queue.push(chunk);
      wake();
    });
    s.on('end', () => {
      done = true;
      wake();
    });
    s.on('error', (err: unknown) => {
      error = err;
      done = true;
      wake();
    });

    if (typeof s.resume === 'function') {
      s.resume();
    }

    while (!done || queue.length > 0) {
      if (queue.length === 0) {
        await new Promise<void>(resolve => {
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
