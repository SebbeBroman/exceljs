import fs from 'fs';

// useful stuff
// eslint-disable-next-line @typescript-eslint/no-unsafe-function-type
type AnyCtor = Function & {super_?: AnyCtor; prototype: object};

const inherits = function (
  cls: AnyCtor,
  superCtor: AnyCtor,
  statics?: Record<string, unknown> | null,
  prototype?: Record<string, unknown> | null,
): void {
  cls.super_ = superCtor;

  if (!prototype) {
    prototype = statics;
    statics = null;
  }

  if (statics) {
    Object.keys(statics).forEach(i => {
      Object.defineProperty(cls, i, Object.getOwnPropertyDescriptor(statics!, i)!);
    });
  }

  const properties: PropertyDescriptorMap = {
    constructor: {
      value: cls,
      enumerable: false,
      writable: false,
      configurable: true,
    },
  };
  if (prototype) {
    Object.keys(prototype).forEach(i => {
      properties[i] = Object.getOwnPropertyDescriptor(prototype!, i)!;
    });
  }

  cls.prototype = Object.create(superCtor.prototype, properties);
};

// oxlint-disable-next-line no-control-regex
const xmlDecodeRegex = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;
const utils = {
  nop(): void {},
  promiseImmediate<T>(value?: T): Promise<T | undefined> {
    return new Promise(resolve => {
      const g = globalThis as typeof globalThis & {setImmediate?: typeof setImmediate};
      if (typeof g.setImmediate === 'function') {
        g.setImmediate(() => {
          resolve(value);
        });
      } else {
        // poorman's setImmediate - must wait at least 1ms
        setTimeout(() => {
          resolve(value);
        }, 1);
      }
    });
  },
  inherits,
  dateToExcel(d: Date, date1904?: boolean): number {
    return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
  },
  excelToDate(v: number, date1904?: boolean): Date {
    const millisecondSinceEpoch = Math.round(
      (v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000,
    );
    return new Date(millisecondSinceEpoch);
  },
  parsePath(filepath: string): {path: string; name: string} {
    const last = filepath.lastIndexOf('/');
    return {
      path: filepath.substring(0, last),
      name: filepath.substring(last + 1),
    };
  },
  getRelsPath(filepath: string): string {
    const path = utils.parsePath(filepath);
    return `${path.path}/_rels/${path.name}.rels`;
  },
  xmlEncode(text: string): string {
    const regexResult = xmlDecodeRegex.exec(text);
    if (!regexResult) return text;

    let result = '';
    let escape = '';
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < text.length; i++) {
      const charCode = text.charCodeAt(i);
      switch (charCode) {
        case 34: // "
          escape = '&quot;';
          break;
        case 38: // &
          escape = '&amp;';
          break;
        case 39: // '
          escape = '&apos;';
          break;
        case 60: // <
          escape = '&lt;';
          break;
        case 62: // >
          escape = '&gt;';
          break;
        case 127:
          escape = '';
          break;
        default: {
          if (charCode <= 31 && (charCode <= 8 || (charCode >= 11 && charCode !== 13))) {
            escape = '';
            break;
          }
          continue;
        }
      }
      if (lastIndex !== i) result += text.substring(lastIndex, i);
      lastIndex = i + 1;
      if (escape) result += escape;
    }
    if (lastIndex !== i) return result + text.substring(lastIndex, i);
    return result;
  },
  xmlDecode(text: string): string {
    return text.replace(/&([a-z]*);/g, c => {
      switch (c) {
        case '&lt;':
          return '<';
        case '&gt;':
          return '>';
        case '&amp;':
          return '&';
        case '&apos;':
          return "'";
        case '&quot;':
          return '"';
        default:
          return c;
      }
    });
  },
  validInt(value: unknown): number {
    const i = parseInt(String(value), 10);
    return !Number.isNaN(i) ? i : 0;
  },

  isDateFmt(fmt: string | null | undefined): boolean {
    if (!fmt) {
      return false;
    }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/\[[^\]]*]/g, '');
    fmt = fmt.replace(/"[^"]*"/g, '');
    // then check for date formatting chars
    const result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
  },

  fs: {
    exists(path: string): Promise<boolean> {
      return new Promise(resolve => {
        fs.access(path, fs.constants.F_OK, err => {
          resolve(!err);
        });
      });
    },
  },

  toIsoDateString(dt: Date): string {
    // Note: original uses toIsoString (typo) and subsstr (typo) — keep runtime behavior
    const s = (dt as Date & {toIsoString(): string}).toIsoString() as string & {
      subsstr(from: number, length?: number): string;
    };
    return s.subsstr(0, 10);
  },

  parseBoolean(value: unknown): boolean {
    return value === true || value === 'true' || value === 1 || value === '1';
  },

  *range(start: number, stop: number, step = 1): Generator<number, void, unknown> {
    const compareOrder = step > 0 ? (a: number, b: number) => a < b : (a: number, b: number) => a > b;
    for (let value = start; compareOrder(value, stop); value += step) {
      yield value;
    }
  },

  toSortedArray(values: Iterable<unknown>): unknown[] {
    const result = Array.from(values);

    // Note: per default, `Array.prototype.sort()` converts values
    // to strings when comparing. Here, if we have numbers, we use
    // numeric sort.
    if (result.every(item => Number.isFinite(item))) {
      const compareNumbers = (a: unknown, b: unknown) => (a as number) - (b as number);
      return result.sort(compareNumbers);
    }

    return result.sort();
  },

  objectFromProps(props: string[], value: unknown = null): Record<string, unknown> {
    // *Note*: Using `reduce` as `Object.fromEntries` requires Node 12+;
    // ExcelJs is >=8.3.0 (as of 2023-10-08).
    // return Object.fromEntries(props.map(property => [property, value]));
    return props.reduce<Record<string, unknown>>((result, property) => {
      result[property] = value;
      return result;
    }, {});
  },
};

export default utils;
export {utils};

// Named re-exports (avoid re-declaring locals like `inherits`)
export const nop = utils.nop;
export const promiseImmediate = utils.promiseImmediate;
export const dateToExcel = utils.dateToExcel;
export const excelToDate = utils.excelToDate;
export const parsePath = utils.parsePath;
export const getRelsPath = utils.getRelsPath;
export const xmlEncode = utils.xmlEncode;
export const xmlDecode = utils.xmlDecode;
export const validInt = utils.validInt;
export const isDateFmt = utils.isDateFmt;
export const toIsoDateString = utils.toIsoDateString;
export const parseBoolean = utils.parseBoolean;
export const range = utils.range;
export const toSortedArray = utils.toSortedArray;
export const objectFromProps = utils.objectFromProps;
export {inherits};
