import type XformState from './xform-state.js';
import BaseXform from './base-xform.js';
import type {XmlStreamLike, XformOptions} from './base-xform.js';

export interface ListXformOptions {
  tag: string;
  always?: boolean;
  count?: boolean;
  empty?: boolean;
  $count?: string;
  $?: Record<string, unknown>;
  childXform: XformState;
  maxItems?: number;
}

class ListXform extends BaseXform<unknown[]> {
  always: boolean;
  count: boolean | undefined;
  empty: boolean | undefined;
  $count: string;
  $: Record<string, unknown> | undefined;
  childXform: XformState;
  maxItems: number | undefined;

  constructor(options: ListXformOptions) {
    super();

    this.tag = options.tag;
    this.always = !!options.always;
    this.count = options.count;
    this.empty = options.empty;
    this.$count = options.$count || 'count';
    this.$ = options.$;
    this.childXform = options.childXform;
    this.maxItems = options.maxItems;
  }

  override prepare(model?: unknown[] | null, options?: XformOptions): void {
    const {childXform} = this;
    if (model) {
      model.forEach((childModel, index) => {
        if (options) {
          options.index = index;
        }
        childXform.prepare(childModel, options);
      });
    }
  }

  override render(xmlStream: XmlStreamLike, model?: unknown[] | null): void {
    if (this.always || (model && model.length)) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, (model && model.length) || 0);
      }

      const {childXform} = this;
      (model || []).forEach((childModel, index) => {
        childXform.render(xmlStream, childModel, index);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  }
}

export default ListXform;
export {ListXform};
