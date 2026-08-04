/**
 * Deep clone used by xform test helpers.
 * When `preserveUndefined` is false, properties whose value is `undefined` are omitted.
 */
export function cloneDeep<T>(obj: T, preserveUndefined = true): T {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }
  if (obj instanceof Date) {
    return obj;
  }
  if (Array.isArray(obj)) {
    const clone: unknown[] = [];
    obj.forEach((value, index) => {
      if (value !== undefined) {
        clone[index] = cloneDeep(value, preserveUndefined);
      } else if (preserveUndefined) {
        clone[index] = undefined;
      }
    });
    return clone as T;
  }
  const clone: Record<string, unknown> = {};
  for (const [name, value] of Object.entries(obj as Record<string, unknown>)) {
    if (value !== undefined) {
      clone[name] = cloneDeep(value, preserveUndefined);
    } else if (preserveUndefined) {
      clone[name] = undefined;
    }
  }
  return clone as T;
}
