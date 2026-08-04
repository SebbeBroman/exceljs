import {Parser} from 'saxen';
import {bufferToString} from './browser-buffer-decode.js';
import {fromReadable} from './async-iterator.js';

/** Normalized SAX event shapes (matches former saxes adapter). */
export type SaxEvent =
  | {eventType: 'opentag'; value: {name: string; attributes: Record<string, string>}}
  | {eventType: 'text'; value: string}
  | {eventType: 'closetag'; value: {name: string}};

/**
 * Stream XML into exceljs-shaped SAX events.
 * Normalized shape (matches former saxes adapter):
 *   {eventType: 'opentag', value: {name, attributes}}
 *   {eventType: 'text', value: string}
 *   {eventType: 'closetag', value: {name}}
 */
export default async function* parseSax(
  iterable: unknown,
): AsyncGenerator<SaxEvent[], void, unknown> {
  // Accept async iterables, arrays, strings, or legacy Node streams — no PassThrough.
  const source = fromReadable(iterable);
  const parser = new Parser();
  let error: Error | undefined;
  let events: SaxEvent[] = [];

  // saxen reports some well-formedness issues (e.g. text outside root) as
  // recoverable `warn`s; exceljs treats them as hard errors like saxes did.
  const fail = (err: unknown): void => {
    if (!error) {
      error = err instanceof Error ? err : new Error(String(err));
    }
  };
  parser.on('error', fail);
  parser.on('warn', fail);

  parser.on('openTag', (name, getAttrs, decodeEntities) => {
    const raw = getAttrs();
    const attributes = Object.create(null) as Record<string, string>;
    for (const key in raw) {
      const v = raw[key];
      // Hot path: cell attrs (r/t/s) almost never contain entities — skip decode.
      attributes[key] = v.indexOf('&') === -1 ? v : decodeEntities(v);
    }
    events.push({eventType: 'opentag', value: {name, attributes}});
  });

  parser.on('text', (value, decodeEntities) => {
    // Shared-string / inline text often has no entities either.
    events.push({
      eventType: 'text',
      value: value.indexOf('&') === -1 ? value : decodeEntities(value),
    });
  });

  parser.on('closeTag', name => {
    events.push({eventType: 'closetag', value: {name}});
  });

  for await (const chunk of source) {
    parser.write(bufferToString(chunk));
    // saxen callbacks are synchronous during write, so all events for this
    // chunk are already in `events` before we reach here.
    if (error) throw error;
    yield events;
    events = [];
  }

  parser.end();
  if (error) throw error;
  if (events.length) {
    yield events;
  }
}
