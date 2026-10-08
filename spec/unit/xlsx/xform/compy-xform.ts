import BaseXform from '../../../../lib/xlsx/xform/base-xform.js';
import type {XmlNode, XmlStreamLike} from '../../../../lib/xlsx/xform/base-xform.js';

interface Child {
  name?: string;
  tag?: string;
  xform: BaseXform;
}

/** Test composite that checks transforms nested inside another XML element. */
class CompyXform extends BaseXform<Record<string, unknown>> {
  override tag: string;
  private attrs?: Record<string, unknown>;
  private children: Child[];
  override map: Record<string, Child>;
  override parser: Child | undefined = undefined;

  constructor(options: {tag: string; attrs?: Record<string, unknown>; children: Child[]}) {
    super();
    this.tag = options.tag;
    this.attrs = options.attrs;
    this.children = options.children;
    this.map = {};
    for (const child of this.children) {
      child.name ??= child.tag;
      child.tag ??= child.name;
      this.map[child.tag!] = child;
    }
  }

  override prepare(model, options): void {
    for (const child of this.children) child.xform.prepare(model[child.tag!], options);
  }

  override render(xmlStream: XmlStreamLike, model): void {
    xmlStream.openNode(this.tag, this.attrs);
    for (const child of this.children) child.xform.render(xmlStream, model[child.name!]);
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.xform.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.model = {};
      return true;
    }
    this.parser = this.map[node.name];
    if (!this.parser) return false;
    this.parser.xform.parseOpen(node);
    return true;
  }

  override parseText(text: string): void {
    this.parser?.xform.parseText(text);
  }

  override parseClose(name: string): boolean {
    if (!this.parser) return false;
    if (!this.parser.xform.parseClose(name)) {
      this.model![this.parser.name!] = this.parser.xform.model;
      this.parser = undefined;
    }
    return true;
  }

  override reconcile(model, options): void {
    for (const child of this.children) child.xform.prepare(model[child.tag!], options);
  }
}
export default CompyXform;
