/**
 * Environment helpers without touching Node's `process` global.
 * (Bare `process` breaks browser ESM; we never want consumers to polyfill it.)
 */

interface ProcessLike {
  versions?: {node?: string} | null;
}

/** True when running under Node.js (not browser / not Deno-without-node-compat). */
export function isNode(): boolean {
  return (
    typeof globalThis !== 'undefined' &&
    (globalThis as {process?: ProcessLike | null}).process != null &&
    typeof (globalThis as {process?: ProcessLike}).process!.versions === 'object' &&
    (globalThis as {process?: ProcessLike}).process!.versions != null &&
    (globalThis as {process?: ProcessLike}).process!.versions!.node != null
  );
}

export function isBrowser(): boolean {
  return !isNode();
}

/** schedule async continuation without process.nextTick */
export function nextTick(fn: () => void): void {
  if (typeof queueMicrotask === 'function') {
    queueMicrotask(fn);
  } else {
    Promise.resolve().then(fn);
  }
}
