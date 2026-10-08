import Encryptor from '../utils/encryptor.js';
import type {ProtectOptions, SheetProtection} from '../model/types.js';

/** Prepare sheet protection explicitly; crypto is included only through this entry. */
export function sheetProtection(password?: string, options?: ProtectOptions): SheetProtection {
  const model: SheetProtection = {sheet: true};
  const spinCount =
    options?.spinCount == null
      ? 100000
      : Number.isFinite(options.spinCount)
        ? Math.round(Math.max(0, options.spinCount))
        : 100000;
  if (password) {
    const salt = Encryptor.randomBytesBase64(16);
    Object.assign(model, {
      algorithmName: 'SHA-512',
      saltValue: salt,
      spinCount,
      hashValue: Encryptor.convertPasswordToHash(password, 'SHA512', salt, spinCount),
    });
  }
  Object.assign(model, options);
  if (password) model.spinCount = spinCount;
  else delete model.spinCount;
  return model;
}
