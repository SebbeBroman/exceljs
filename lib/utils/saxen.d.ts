/**
 * Minimal types for saxen (package ships no TypeScript declarations).
 * Scoped to exceljs's usage in parse-sax.ts.
 */
declare module 'saxen' {
  export class Parser {
    on(event: 'error' | 'warn', handler: (err: unknown) => void): void;
    on(
      event: 'openTag',
      handler: (
        name: string,
        getAttrs: () => Record<string, string>,
        decodeEntities: (s: string) => string,
      ) => void,
    ): void;
    on(
      event: 'text',
      handler: (value: string, decodeEntities: (s: string) => string) => void,
    ): void;
    on(event: 'closeTag', handler: (name: string) => void): void;
    write(chunk: string): void;
    end(): void;
  }
}
