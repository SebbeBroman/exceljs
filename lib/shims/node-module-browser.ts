/**
 * Browser stub for node:module — createRequire is unavailable.
 * Doc features use dynamic import() in the browser instead.
 */
export function createRequire(): () => never {
  return function browserRequire(): never {
    throw new Error('createRequire is not available in the browser');
  };
}

const nodeModuleBrowser = {createRequire};
export default nodeModuleBrowser;
