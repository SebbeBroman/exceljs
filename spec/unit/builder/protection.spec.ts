import {describe, it, expect} from 'vite-plus/test';
import {sheetProtection} from '../../../protection.js';
import {workbook, writeBuffer, load} from '../../../excel.js';
import Encryptor from '../../../lib/utils/encryptor.js';

describe('opt-in sheet protection', () => {
  it('prepares a verifiable hash and preserves it through plain models and edit loops', async () => {
    const model = sheetProtection('pw', {spinCount: 2, selectLockedCells: false});
    expect(model.hashValue).toBe(
      Encryptor.convertPasswordToHash('pw', 'SHA512', model.saltValue!, 2),
    );
    const plain = workbook().sheet('Locked').row([1]).protect(model).build();
    const loaded = await load(await writeBuffer(plain));
    expect(loaded.sheets[0]!.sheetProtection).toMatchObject(model);
    const edited = await load(await workbook(loaded).sheet('Locked').cell('A2', 2).writeBuffer());
    expect(edited.sheets[0]!.sheetProtection).toEqual(loaded.sheets[0]!.sheetProtection);
  });
  it('normalizes spin counts and excludes hashing fields without a password', () => {
    expect(sheetProtection('pw', {spinCount: -1}).spinCount).toBe(0);
    expect(sheetProtection('pw', {spinCount: 1.4}).spinCount).toBe(1);
    expect(sheetProtection(undefined, {spinCount: 2, selectLockedCells: false})).toEqual({
      sheet: true,
      selectLockedCells: false,
    });
  });
  it('rejects the previous password API with an actionable error', () => {
    expect(() =>
      workbook()
        .sheet('Old')
        .protect('pw' as never),
    ).toThrow(/protection model/);
  });
});
