const {toString} = Object.prototype;

function isUndefined(val: unknown): val is undefined {
  return toString.call(val) === '[object Undefined]';
}

function isObject(val: unknown): val is Record<string, unknown> {
  return toString.call(val) === '[object Object]';
}

/** Deep equality for plain objects, arrays, and primitives. */
export function isEqual(a: unknown, b: unknown): boolean {
  const aType = typeof a;
  const bType = typeof b;
  const aArray = Array.isArray(a);
  const bArray = Array.isArray(b);

  if (aType !== bType) {
    return false;
  }
  switch (typeof a) {
    case 'object': {
      if (aArray || bArray) {
        if (aArray && bArray) {
          return (
            (a as unknown[]).length === (b as unknown[]).length &&
            (a as unknown[]).every((aValue, index) => isEqual(aValue, (b as unknown[])[index]))
          );
        }
        return false;
      }

      if (a === null || b === null) {
        return a === b;
      }

      const keys = Object.keys(a as object);
      if (Object.keys(b as object).length !== keys.length) {
        return false;
      }

      for (const key of keys) {
        if (!Object.hasOwn(b as object, key)) {
          return false;
        }
      }

      return keys.every(key =>
        isEqual((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key]),
      );
    }
    default:
      return a === b;
  }
}

/**
 * Deep-merge sources into the first argument (mutates and returns it).
 * Arrays and plain objects are merged recursively; undefined source values are skipped.
 */
export function deepMerge(...args: unknown[]): Record<string, unknown> {
  const target = (args[0] || {}) as Record<string, unknown>;
  const {length} = args;

  function assignValue(val: unknown, key: string): void {
    const src = target[key];
    const copyIsArray = Array.isArray(val);
    if (isObject(val) || copyIsArray) {
      let clone: unknown;
      if (copyIsArray) {
        clone = src && Array.isArray(src) ? src : [];
      } else {
        clone = src && isObject(src) ? src : {};
      }
      target[key] = deepMerge(clone, val);
    } else if (!isUndefined(val)) {
      target[key] = val;
    }
  }

  for (let i = 0; i < length; i++) {
    const source = args[i];
    if (source) {
      if (Array.isArray(source)) {
        (source as unknown[]).forEach((val, index) => {
          assignValue(val, String(index));
        });
      } else if (typeof source === 'object') {
        for (const key of Object.keys(source as object)) {
          assignValue((source as Record<string, unknown>)[key], key);
        }
      }
    }
  }
  return target;
}
