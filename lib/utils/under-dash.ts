const {toString} = Object.prototype;
const escapeHtmlRegex = /["&<>]/;

type EachCallback<T> = (value: T, key: string | number) => void;
type SomeCallback<T> = (value: T, key: string | number) => unknown;
type MapCallback<T, R> = (value: T, key: string | number) => R;

const _ = {
  each: function each<T>(
    obj: T[] | Record<string, T> | null | undefined,
    cb: EachCallback<T>,
  ): void {
    if (obj) {
      if (Array.isArray(obj)) {
        obj.forEach(cb as (value: T, index: number, array: T[]) => void);
      } else {
        Object.keys(obj).forEach(key => {
          cb(obj[key], key);
        });
      }
    }
  },

  some: function some<T>(
    obj: T[] | Record<string, T> | null | undefined,
    cb: SomeCallback<T>,
  ): boolean {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.some(cb as (value: T, index: number, array: T[]) => unknown) as boolean;
      }
      return Object.keys(obj).some(key => cb(obj[key], key));
    }
    return false;
  },

  every: function every<T>(
    obj: T[] | Record<string, T> | null | undefined,
    cb: SomeCallback<T>,
  ): boolean {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.every(cb as (value: T, index: number, array: T[]) => unknown) as boolean;
      }
      return Object.keys(obj).every(key => cb(obj[key], key));
    }
    return true;
  },

  map: function map<T, R>(
    obj: T[] | Record<string, T> | null | undefined,
    cb: MapCallback<T, R>,
  ): R[] {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.map(cb as (value: T, index: number, array: T[]) => R);
      }
      return Object.keys(obj).map(key => cb(obj[key], key));
    }
    return [];
  },

  keyBy<T extends Record<string, unknown>>(a: T[], p: keyof T & string): Record<string, T> {
    return a.reduce<Record<string, T>>((o, v) => {
      o[String(v[p])] = v;
      return o;
    }, {});
  },

  isEqual: function isEqual(a: unknown, b: unknown): boolean {
    const aType = typeof a;
    const bType = typeof b;
    const aArray = Array.isArray(a);
    const bArray = Array.isArray(b);
    let keys: string[];

    if (aType !== bType) {
      return false;
    }
    switch (typeof a) {
      case 'object':
        if (aArray || bArray) {
          if (aArray && bArray) {
            return (
              (a as unknown[]).length === (b as unknown[]).length &&
              (a as unknown[]).every((aValue, index) => {
                const bValue = (b as unknown[])[index];
                return _.isEqual(aValue, bValue);
              })
            );
          }
          return false;
        }

        if (a === null || b === null) {
          return a === b;
        }

        // Compare object keys and values
        keys = Object.keys(a as object);

        if (Object.keys(b as object).length !== keys.length) {
          return false;
        }

        for (const key of keys) {
          if (!Object.hasOwn(b as object, key)) {
            return false;
          }
        }

        return _.every(a as Record<string, unknown>, (aValue, key) => {
          const bValue = (b as Record<string, unknown>)[key as string];
          return _.isEqual(aValue, bValue);
        });

      default:
        return a === b;
    }
  },

  escapeHtml(html: string): string {
    const regexResult = escapeHtmlRegex.exec(html);
    if (!regexResult) return html;

    let result = '';
    let escape = '';
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < html.length; i++) {
      switch (html.charAt(i)) {
        case '"':
          escape = '&quot;';
          break;
        case '&':
          escape = '&amp;';
          break;
        case "'":
          escape = '&apos;';
          break;
        case '<':
          escape = '&lt;';
          break;
        case '>':
          escape = '&gt;';
          break;
        default:
          continue;
      }
      if (lastIndex !== i) result += html.substring(lastIndex, i);
      lastIndex = i + 1;
      result += escape;
    }
    if (lastIndex !== i) return result + html.substring(lastIndex, i);
    return result;
  },

  strcmp(a: string, b: string): number {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  },

  isUndefined(val: unknown): val is undefined {
    return toString.call(val) === '[object Undefined]';
  },

  isObject(val: unknown): val is Record<string, unknown> {
    return toString.call(val) === '[object Object]';
  },

  deepMerge(...args: unknown[]): Record<string, unknown> {
    const target = (args[0] || {}) as Record<string, unknown>;
    const {length} = args;

    let src: unknown;
    let clone: unknown;
    let copyIsArray: boolean;

    function assignValue(val: unknown, key: string | number): void {
      src = target[key as string];
      copyIsArray = Array.isArray(val);
      if (_.isObject(val) || copyIsArray) {
        if (copyIsArray) {
          copyIsArray = false;
          clone = src && Array.isArray(src) ? src : [];
        } else {
          clone = src && _.isObject(src) ? src : {};
        }
        target[key as string] = _.deepMerge(clone, val);
      } else if (!_.isUndefined(val)) {
        target[key as string] = val;
      }
    }

    for (let i = 0; i < length; i++) {
      _.each(args[i] as Record<string, unknown>, assignValue);
    }
    return target;
  },
};

export default _;
export {_};
