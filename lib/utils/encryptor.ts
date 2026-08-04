import {sha512} from '@noble/hashes/sha2.js';
import {
  concat,
  encodeBase64,
  from as bytesFrom,
  toPublic,
  utf16leEncode,
  writeUInt32LE,
} from './bytes.js';

/** Normalize OOXML / Node-style names to a supported algorithm id. */
function normalizeAlgorithm(hashAlgorithm: string): 'sha512' {
  const key = hashAlgorithm.toLowerCase().replace(/-/g, '');
  if (key === 'sha512') {
    return 'sha512';
  }
  throw new Error(`Hash algorithm '${hashAlgorithm}' not supported!`);
}

/**
 * SHA-512 over concatenated buffers (Excel sheet-protection hash primitive).
 * Pure JS via @noble/hashes — works in Node and browsers without a crypto shim.
 */
function digestSha512(...buffers: Uint8Array[]): Uint8Array {
  if (buffers.length === 1) {
    return sha512(buffers[0]!);
  }
  return sha512(concat(buffers));
}

const Encryptor = {
  /**
   * Calculate a hash of the concatenated buffers with the given algorithm.
   * @param algorithm - The hash algorithm (only SHA-512 / sha512 supported).
   * @returns The hash bytes
   */
  hash(algorithm: string, ...buffers: Uint8Array[]): Uint8Array {
    normalizeAlgorithm(algorithm);
    return digestSha512(...buffers);
  },
  /**
   * Convert a password into an encryption key (Excel sheet protection spin-hash).
   * @param password - The password
   * @param hashAlgorithm - The hash algorithm (SHA-512 only)
   * @param saltValue - The salt value (base64)
   * @param spinCount - The spin count
   * @returns The hash as base64
   */
  convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number,
  ): string {
    normalizeAlgorithm(hashAlgorithm);

    // Password must be in unicode buffer
    const passwordBuffer = utf16leEncode(password);
    // Generate the initial hash
    let key = digestSha512(bytesFrom(saltValue, 'base64'), passwordBuffer);

    // Reuse one 68-byte block: 64-byte digest + LE32 iterator (avoids allocs in the loop)
    const block = new Uint8Array(64 + 4);
    for (let i = 0; i < spinCount; i++) {
      block.set(key, 0);
      // this is the 'special' element of Excel password hashing
      // that stops us from using crypto.pbkdf2()
      writeUInt32LE(block, i, 64);
      key = sha512(block);
    }
    return encodeBase64(key);
  },
  /**
   * Generates cryptographically strong pseudo-random data (Web Crypto getRandomValues).
   * @param size The size argument is a number indicating the number of bytes to generate.
   */
  randomBytes(size: number): Uint8Array {
    const u8 = new Uint8Array(size);
    globalThis.crypto.getRandomValues(u8);
    return toPublic(u8) as Uint8Array;
  },
  /**
   * Base64-encode random bytes (portable; does not require Node Buffer).
   */
  randomBytesBase64(size: number): string {
    return encodeBase64(this.randomBytes(size));
  },
};
export default Encryptor;
export {Encryptor};
