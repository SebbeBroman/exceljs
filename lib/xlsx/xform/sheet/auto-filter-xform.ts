import colCache from '../../../utils/col-cache.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface AutoFilterAddress {
  row: number;
  column: number;
}

export interface AutoFilterRangeModel {
  from: string | AutoFilterAddress;
  to: string | AutoFilterAddress;
}

export type AutoFilterModel = string | AutoFilterRangeModel;

class AutoFilterXform extends BaseXform<string> {
  override tag = 'autoFilter';

  override render(xmlStream: XmlStreamLike, model?: AutoFilterModel | null): void {
    if (model) {
      if (typeof model === 'string') {
        // assume range
        xmlStream.leafNode('autoFilter', {ref: model});
      } else {
        const getAddress = function (addr: string | AutoFilterAddress): string {
          if (typeof addr === 'string') {
            return addr;
          }
          return colCache.getAddress(addr.row, addr.column).address;
        };

        const firstAddress = getAddress(model.from);
        const secondAddress = getAddress(model.to);
        if (firstAddress && secondAddress) {
          xmlStream.leafNode('autoFilter', {ref: `${firstAddress}:${secondAddress}`});
        }
      }
    }
  }
}

export default AutoFilterXform;
export {AutoFilterXform};
