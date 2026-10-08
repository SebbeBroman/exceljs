import type {ProtectOptions, SheetProtection} from './excel.js';
export type {ProtectOptions, SheetProtection} from './excel.js';
/** Prepare a protection model synchronously; SHA-512 is isolated to this entry. */
export declare function sheetProtection(
  password?: string,
  options?: ProtectOptions,
): SheetProtection;
