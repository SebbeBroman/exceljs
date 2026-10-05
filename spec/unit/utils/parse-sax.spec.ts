import {describe, it, expect} from 'vite-plus/test';
import parseSax from '../../../lib/utils/parse-sax.ts';

async function drain(input: unknown): Promise<unknown[]> {
  const out: unknown[] = [];
  for await (const events of parseSax(input) as AsyncGenerator<unknown[]>) {
    out.push(...events);
  }
  return out;
}

describe('parse-sax saxen adapter', () => {
  it('parses open/text/close events', async () => {
    const events = await drain(['<root><a b="1">hi</a></root>']);
    expect(events.length).toBeGreaterThan(0);
  });

  it('treats well-formedness warns as errors (saxes parity)', async () => {
    // Text outside the root element triggers saxen `warn`; exceljs treats it as fatal.
    await expect(drain(['oops<root/>'])).rejects.toThrow();
  });

  it('decodes entities only when needed', async () => {
    const events = (await drain(['<root><a>fish &amp; chips</a></root>'])) as Array<{
      eventType: string;
      value: unknown;
    }>;
    const text = events.find(e => e.eventType === 'text');
    expect(text?.value).toEqual('fish & chips');
  });
});
