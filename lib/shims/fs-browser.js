/**
 * Browser stub for Node's `fs`.
 * File path APIs are unavailable in the browser — use buffer/stream APIs instead:
 *   await workbook.xlsx.load(arrayBuffer)
 *   await workbook.xlsx.writeBuffer()
 */

function unavailable(method) {
  return () => {
    throw new Error(
      `fs.${method}() is not available in the browser. Use workbook.xlsx.load() / writeBuffer() (or read/write with Blob streams) instead of *File path APIs.`,
    );
  };
}

export function readFile() {
  return unavailable('readFile')();
}
export function writeFile() {
  return unavailable('writeFile')();
}
export function createReadStream() {
  return unavailable('createReadStream')();
}
export function createWriteStream() {
  return unavailable('createWriteStream')();
}
export function access(_path, _mode, cb) {
  const err = new Error('fs.access is not available in the browser');
  if (typeof cb === 'function') cb(err);
  else return Promise.reject(err);
}
export const constants = {F_OK: 0, R_OK: 4, W_OK: 2, X_OK: 1};

export default {
  readFile,
  writeFile,
  createReadStream,
  createWriteStream,
  access,
  constants,
};
