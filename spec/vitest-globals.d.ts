/**
 * Ambient globals for the Vitest suite (globals: true + historical helpers).
 */

declare global {
  /**
   * Historical sync loader for lib modules (see spec/utils/verquire.ts).
   * Returns the module default export (CJS-compatible unwrapping).
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function verquire(modPath: string): any;

  // Mocha aliases used by migrated specs (Vitest globals also provide describe/it/…)
  const context: typeof describe;
}

export {};
