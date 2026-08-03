/**
 * Browser stub for node:module — createRequire is unavailable.
 * Doc features use dynamic import() in the browser instead.
 */
export function createRequire() {
  return function browserRequire() {
    throw new Error('createRequire is not available in the browser');
  };
}

export default {createRequire};
