import {SaxesParser} from 'saxes';
import {bufferToString} from './browser-buffer-decode.js';
import {fromReadable} from './async-iterator.js';

export default async function* parseSax(iterable) {
  // Accept async iterables, arrays, strings, or legacy Node streams — no PassThrough.
  const source = fromReadable(iterable);
  const saxesParser = new SaxesParser();
  let error;
  saxesParser.on('error', err => {
    error = err;
  });
  let events = [];
  saxesParser.on('opentag', value => events.push({eventType: 'opentag', value}));
  saxesParser.on('text', value => events.push({eventType: 'text', value}));
  saxesParser.on('closetag', value => events.push({eventType: 'closetag', value}));
  for await (const chunk of source) {
    saxesParser.write(bufferToString(chunk));
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    if (error) throw error;
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue
    yield events;
    events = [];
  }
}
