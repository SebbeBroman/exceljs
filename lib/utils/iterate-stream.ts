import type {EventEmitter} from 'node:events';
import type {Readable} from 'node:stream';

/** Minimal event-emitter surface used by iterateStream. */
interface StreamLike {
  on(event: string, listener: (...args: unknown[]) => void): unknown;
  resume?(): unknown;
  pause?(): unknown;
  addListener?(event: string, listener: (...args: unknown[]) => void): unknown;
  removeListener?(event: string, listener: (...args: unknown[]) => void): unknown;
}

export default async function* iterateStream(
  stream: StreamLike | Readable,
): AsyncGenerator<unknown, void, unknown> {
  const contents: unknown[] = [];
  stream.on('data', (data: unknown) => contents.push(data));

  let resolveStreamEndedPromise!: () => void;
  const streamEndedPromise = new Promise<void>(resolve => (resolveStreamEndedPromise = resolve));

  let ended = false;
  stream.on('end', () => {
    ended = true;
    resolveStreamEndedPromise();
  });

  let error: unknown = false;
  stream.on('error', (err: unknown) => {
    error = err;
    resolveStreamEndedPromise();
  });

  while (!ended || contents.length > 0) {
    if (contents.length === 0) {
      stream.resume?.();

      await Promise.race([once(stream, 'data'), streamEndedPromise]);
    } else {
      stream.pause?.();
      const data = contents.shift();
      yield data;
    }
    if (error) throw error;
  }
  resolveStreamEndedPromise();
}

function once(eventEmitter: StreamLike | EventEmitter, type: string): Promise<void> {
  // TODO: Use __esm_0.once when node v10 is dropped
  return new Promise(resolve => {
    let fired = false;
    const handler = (): void => {
      if (!fired) {
        fired = true;
        if (typeof eventEmitter.removeListener === 'function') {
          eventEmitter.removeListener(type, handler);
        }
        resolve();
      }
    };
    if (typeof eventEmitter.addListener === 'function') {
      eventEmitter.addListener(type, handler);
    } else {
      eventEmitter.on(type, handler);
    }
  });
}
