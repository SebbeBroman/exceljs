import BaseXform from './base-xform.js';
import type {XmlStreamLike, XmlNode, XformOptions} from './base-xform.js';

export interface ListXformOptions {
  tag: string;
  always?: boolean;
  count?: boolean;
  empty?: boolean;
  $count?: string;
  $?: Record<string, unknown>;
  childXform: BaseXform;
  maxItems?: number;
}

class ListXform extends BaseXform<unknown[]> {
  always: boolean;
  count: boolean | undefined;
  empty: boolean | undefined;
  $count: string;
  $: Record<string, unknown> | undefined;
  childXform: BaseXform;
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
