import type {ListXformOptions} from '../xform/list-xform.js';
export type {ListXformOptions} from '../xform/list-xform.js';
import type XformState from '../xform/xform-state.js';
import BaseXform from '../base-parser.js';
import type {XmlNode, XformOptions} from '../base-parser.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = [];
        return true;
      default:
        if (this.childXform.parseOpen(node)) {
          this.parser = this.childXform;
          return true;
        }
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        (this.model as unknown[]).push(this.parser.model);
        this.parser = undefined;

        if (this.maxItems && (this.model as unknown[]).length > this.maxItems) {
          throw new Error(`Max ${this.childXform.tag} count (${this.maxItems}) exceeded`);
        }
      }
      return true;
    }

    return false;
  }

  override reconcile(model?: unknown[] | null, options?: XformOptions): void {
    if (model) {
      const {childXform} = this;
      model.forEach(childModel => {
        childXform.reconcile(childModel, options);
      });
    }
  }
}

export default ListXform;
export {ListXform};
