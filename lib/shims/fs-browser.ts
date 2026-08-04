/**
 * Browser stub for Node's `fs`.
 * File path APIs are unavailable in the browser — use buffer/stream APIs instead:
 *   await workbook.xlsx.load(arrayBuffer)
 *   await workbook.xlsx.writeBuffer()
 */

function unavailable(method: string): () => never {
  return () => {
    throw new Error(
      `fs.${method}() is not available in the browser. Use workbook.xlsx.load() / writeBuffer() (or read/write with Blob streams) instead of *File path APIs.`,
    );
  };
}

export function readFile(): never {
  return unavailable('readFile')();
}
export function writeFile(): never {
  return unavailable('writeFile')();
}
export function createReadStream(): never {
  return unavailable('createReadStream')();
}
export function createWriteStream(): never {
  return unavailable('createWriteStream')();
}
export function access(
  _path: string,
  _mode?: number | ((err: Error) => void),
  cb?: (err: Error) => void,
): void | Promise<never> {
  const err = new Error('fs.access is not available in the browser');
  if (typeof cb === 'function') {
    cb(err);
    return;
  }
  if (typeof _mode === 'function') {
    _mode(err);
    return;
  }
  return Promise.reject(err);
}
export const constants = {F_OK: 0, R_OK: 4, W_OK: 2, X_OK: 1} as const;

const fsBrowser = {
  readFile,
  writeFile,
  createReadStream,
  createWriteStream,
  access,
  constants,
};

export default fsBrowser;
