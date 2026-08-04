import crypto from 'crypto';
import {concat, from as bytesFrom, utf16leEncode, writeUInt32LE} from './bytes.js';

const Encryptor = {
  /**
   * Calculate a hash of the concatenated buffers with the given algorithm.
   * @param algorithm - The hash algorithm.
   * @returns The hash
   */
  hash(algorithm: string, ...buffers: Uint8Array[]): Buffer {
    const hash = crypto.createHash(algorithm);
    // Node crypto accepts Uint8Array
    hash.update(concat(buffers));
    return hash.digest();
  },
  /**
   * Convert a password into an encryption key
   * @param password - The password
   * @param hashAlgorithm - The hash algoritm
   * @param saltValue - The salt value
   * @param spinCount - The spin count
   * @returns The hash as base64
   */
  convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number,
  ): string {
    hashAlgorithm = hashAlgorithm.toLowerCase();
    const hashes = crypto.getHashes();
    if (hashes.indexOf(hashAlgorithm) < 0) {
      throw new Error(`Hash algorithm '${hashAlgorithm}' not supported!`);
    }

    // Password must be in unicode buffer
    const passwordBuffer = utf16leEncode(password);
    // Generate the initial hash
    let key = this.hash(hashAlgorithm, bytesFrom(saltValue, 'base64'), passwordBuffer);
    // Now regenerate until spin count
    for (let i = 0; i < spinCount; i++) {
      const iterator = new Uint8Array(4);
      // this is the 'special' element of Excel password hashing
      // that stops us from using crypto.pbkdf2()
      writeUInt32LE(iterator, i, 0);
      key = this.hash(hashAlgorithm, key, iterator);
    }
    return key.toString('base64');
  },
  /**
   * Generates cryptographically strong pseudo-random data.
   * @param size The size argument is a number indicating the number of bytes to generate.
   */
  randomBytes(size: number): Buffer {
    return crypto.randomBytes(size);
  },
};
export default Encryptor;
export {Encryptor};
