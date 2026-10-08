import {describe, it, expect} from 'vite-plus/test';
import TwoCellAnchorXform from '../../../../../lib/xlsx/xform/drawing/two-cell-anchor-xform.js';

// Intentionally malformed models exercise defensive reconciliation.
describe('TwoCellAnchorXform', () => {
  describe('reconcile', () => {
    it('should not throw on null picture', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() => twoCell.reconcile({picture: null} as never, {})).not.toThrow();
    });
    it('should not throw on null tl', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() =>
        twoCell.reconcile({br: {col: 1, row: 1}} as never, {})
      ).not.toThrow();
    });
    it('should not throw on null br', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() =>
        twoCell.reconcile({tl: {col: 1, row: 1}} as never, {})
      ).not.toThrow();
    });
  });
});
