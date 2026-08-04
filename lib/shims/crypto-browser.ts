/**
 * Minimal browser stub for Node's `crypto`.
 * - randomBytes → Web Crypto getRandomValues
 * - createHash / password hashing used by worksheet.protect() is NOT supported
 *   client-side (Node crypto hashing is sync; Web Crypto is async).
 */
import {toPublic} from '../utils/bytes.js';

export function getHashes(): string[] {
  return [];
}

export function createHash(algorithm: string): never {
  throw new Error(
    `crypto.createHash('${algorithm}') is not available in the browser build. ` +
      `Avoid worksheet.protect() / password hashing on the client, or run that on the server.`,
  );
}

export function randomBytes(size: number): Uint8Array {
  const u8 = new Uint8Array(size);
  globalThis.crypto.getRandomValues(u8);
  return toPublic(u8) as Uint8Array;
}

const cryptoBrowser = {
  getHashes,
  createHash,
  randomBytes,
};

export default cryptoBrowser;
