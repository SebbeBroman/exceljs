/**
 * Minimal browser stub for Node's `crypto`.
 * - randomBytes → Web Crypto getRandomValues
 * - createHash / password hashing used by worksheet.protect() is NOT supported
 *   client-side (Node crypto hashing is sync; Web Crypto is async).
 */
import {Buffer} from 'buffer';

export function getHashes() {
  return [];
}

export function createHash(algorithm) {
  throw new Error(
    `crypto.createHash('${algorithm}') is not available in the browser build. ` +
      `Avoid worksheet.protect() / password hashing on the client, or run that on the server.`
  );
}

export function randomBytes(size) {
  const u8 = new Uint8Array(size);
  globalThis.crypto.getRandomValues(u8);
  return Buffer.from(u8);
}

export default {
  getHashes,
  createHash,
  randomBytes,
};
