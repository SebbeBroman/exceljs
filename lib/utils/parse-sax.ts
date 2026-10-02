import {Parser} from 'saxen';
import {bufferToString} from './browser-buffer-decode.js';
import {fromReadable} from './async-iterator.js';

/** Normalized SAX event shapes (matches former saxes adapter). */
export type SaxEvent =
  | {eventType: 'opentag'; value: {name: string; attributes: Record<string, string>}}
  | {eventType: 'text'; value: string}
  | {eventType: 'closetag'; value: {name: string}};

/** Look up one attribute. The function is reused across tags; do not store it. */
export type SaxAttributeGetter = (name: string) => string | undefined;

export interface SaxWalkHandlers {
  /**
   * `attr` and `materialize` are reused across tags. Read them before returning.
   * `materialize` copies every attribute into a fresh object (xform path).
   * `attr` decodes a single value and skips saxen's attribute parse until called.
   */
  onOpen(
    name: string,
    attr: SaxAttributeGetter,
    materialize: () => Record<string, string>,
  ): void;
  onText(value: string): void;
  onClose(name: string): void;
}

function decoded(value: string, decodeEntities: (text: string) => string): string {
  return value.indexOf('&') === -1 ? value : decodeEntities(value);
}

interface SaxSession {
  parser: Parser;
  error: () => Error | undefined;
}

function openSax(handlers: SaxWalkHandlers): SaxSession {
  const parser = new Parser();
  let error: Error | undefined;

  // saxen reports some well-formedness issues (e.g. text outside root) as
  // recoverable `warn`s; exceljs treats them as hard errors like saxes did.
  const fail = (err: unknown): void => {
    if (!error) {
      error = err instanceof Error ? err : new Error(String(err));
    }
  };
  parser.on('error', fail);
  parser.on('warn', fail);

  let raw: Record<string, string> | null = null;
  let decodeEntities: (text: string) => string = text => text;
  let getAttrs: () => Record<string, string> | false = () => ({});

  const loadRaw = (): Record<string, string> => {
    if (raw === null) {
      const got = getAttrs();
      raw = got && typeof got === 'object' ? got : {};
    }
    return raw;
  };
  const attr: SaxAttributeGetter = name => {
    const value = loadRaw()[name];
    if (value == null) return undefined;
    return decoded(value, decodeEntities);
  };
  const materialize = (): Record<string, string> => {
    const sourceAttrs = loadRaw();
    const attributes = Object.create(null) as Record<string, string>;
    for (const key in sourceAttrs) {
      attributes[key] = decoded(sourceAttrs[key]!, decodeEntities);
    }
    return attributes;
  };

  parser.on('openTag', (name, nextGetAttrs, nextDecode) => {
    raw = null;
    getAttrs = nextGetAttrs;
    decodeEntities = nextDecode;
    handlers.onOpen(name, attr, materialize);
  });
  parser.on('text', (value, nextDecode) => {
    handlers.onText(decoded(value, nextDecode));
  });
  parser.on('closeTag', name => {
    handlers.onClose(name);
  });

  return {parser, error: () => error};
}

function finishSax(session: SaxSession): void {
  const endError = session.parser.end();
  const error = session.error();
  if (error) throw error;
  if (endError) {
    throw endError instanceof Error ? endError : new Error(String(endError));
  }
}

/**
 * Drive saxen from callbacks, yielding once per input chunk so the caller can
 * release rows before the next chunk is parsed. No per-tag event objects.
 */
export async function* eachSaxChunk(
  iterable: unknown,
  handlers: SaxWalkHandlers,
): AsyncGenerator<void, void, unknown> {
  const source = fromReadable(iterable);
  const session = openSax(handlers);
  for await (const chunk of source) {
    session.parser.write(bufferToString(chunk));
    const error = session.error();
    if (error) throw error;
    yield;
  }
  finishSax(session);
  yield;
}

/**
 * Stream XML into exceljs-shaped SAX events.
 * Xforms retain attribute objects, so each open tag gets its own copy.
 */
export default async function* parseSax(
  iterable: unknown,
): AsyncGenerator<SaxEvent[], void, unknown> {
  const source = fromReadable(iterable);
  const events: SaxEvent[] = [];
  const session = openSax({
    onOpen(name, _attr, materialize) {
      events.push({eventType: 'opentag', value: {name, attributes: materialize()}});
    },
    onText(value) {
      events.push({eventType: 'text', value});
    },
    onClose(name) {
      events.push({eventType: 'closetag', value: {name}});
    },
  });

  for await (const chunk of source) {
    session.parser.write(bufferToString(chunk));
    const error = session.error();
    if (error) throw error;
    if (events.length === 0) continue;
    const batch = events.slice();
    events.length = 0;
    yield batch;
  }

  finishSax(session);
  if (events.length) {
    yield events;
  }
}
