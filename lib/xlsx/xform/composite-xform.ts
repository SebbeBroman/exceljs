import BaseXform from './base-xform.js';
import type {XmlNode} from './base-xform.js';

/* 'virtual' methods used as a form of documentation */

// base class for xforms that are composed of other xforms
// offers some default implementations
class CompositeXform<TModel = Record<string, unknown>> extends BaseXform<TModel> {
  declare map: Record<string, BaseXform>;

  createNewModel(_node?: XmlNode): TModel {
    return {} as TModel;
  }

  override parseOpen(node: XmlNode): boolean {
    // Typical pattern for composite xform
    this.parser = this.parser || this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (node.name === this.tag) {
      this.model = this.createNewModel(node);
      return true;
    }

    return false;
  }

  override parseText(text: string): void {
    // Default implementation. Send text to child parser
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  onParserClose(name: string, parser: BaseXform): void {
    // parseClose has seen a child parser close
    // now need to incorporate into this.model somehow
    (this.model as Record<string, unknown>)[name] = parser.model;
  }

  override parseClose(name?: string): boolean {
    // Default implementation
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.onParserClose(name as string, this.parser);
        this.parser = undefined;
      }
      return true;
    }

    return name !== this.tag;
  }
}

export default CompositeXform;
export {CompositeXform};
