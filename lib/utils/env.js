/**
 * Environment helpers without touching Node's `process` global.
 * (Bare `process` breaks browser ESM; we never want consumers to polyfill it.)
 */

/** True when running under Node.js (not browser / not Deno-without-node-compat). */
export function isNode() {
  return (
    typeof globalThis !== 'undefined' &&
    globalThis.process != null &&
    typeof globalThis.process.versions === 'object' &&
    globalThis.process.versions != null &&
    globalThis.process.versions.node != null
  );
}

export function isBrowser() {
  return !isNode();
}

/** schedule async continuation without process.nextTick */
export function nextTick(fn) {
  if (typeof queueMicrotask === 'function') {
    queueMicrotask(fn);
  } else {
    Promise.resolve().then(fn);
  }
}
