import ColorXform from './color-xform.js';
import BooleanXform from '../simple/boolean-xform.js';
import IntegerXform from '../simple/integer-xform.js';
import StringXform from '../simple/string-xform.js';
import UnderlineXform from './underline-xform.js';
import _ from '../../../utils/under-dash.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface FontXformOptions {
  tagName: string;
  fontNameTag: string;
}

export interface FontMapEntry {
  prop: string;
  xform: BaseXform;
}

// Font encapsulates translation from font model to xlsx
class FontXform extends BaseXform<Record<string, unknown>> {
  options: FontXformOptions;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: Record<string, any>;

  constructor(options?: Partial<FontXformOptions>) {
    super();

    this.options = (options as FontXformOptions) || FontXform.OPTIONS;

    this.map = {
      b: {prop: 'bold', xform: new BooleanXform({tag: 'b', attr: 'val'})},
      i: {prop: 'italic', xform: new BooleanXform({tag: 'i', attr: 'val'})},
      u: {prop: 'underline', xform: new UnderlineXform()},
      charset: {prop: 'charset', xform: new IntegerXform({tag: 'charset', attr: 'val'})},
      color: {prop: 'color', xform: new ColorXform()},
      condense: {prop: 'condense', xform: new BooleanXform({tag: 'condense', attr: 'val'})},
      extend: {prop: 'extend', xform: new BooleanXform({tag: 'extend', attr: 'val'})},
      family: {prop: 'family', xform: new IntegerXform({tag: 'family', attr: 'val'})},
      outline: {prop: 'outline', xform: new BooleanXform({tag: 'outline', attr: 'val'})},
      vertAlign: {prop: 'vertAlign', xform: new StringXform({tag: 'vertAlign', attr: 'val'})},
      scheme: {prop: 'scheme', xform: new StringXform({tag: 'scheme', attr: 'val'})},
      shadow: {prop: 'shadow', xform: new BooleanXform({tag: 'shadow', attr: 'val'})},
      strike: {prop: 'strike', xform: new BooleanXform({tag: 'strike', attr: 'val'})},
      sz: {prop: 'size', xform: new IntegerXform({tag: 'sz', attr: 'val'})},
    };
    this.map[this.options.fontNameTag] = {
      prop: 'name',
      xform: new StringXform({tag: this.options.fontNameTag, attr: 'val'}),
    };
    this.tag = this.options.tagName;
  }

  override render(xmlStream: XmlStreamLike, model?: Record<string, unknown> | null): void {
    const {map} = this;

    xmlStream.openNode(this.options.tagName);
    _.each(this.map, (defn: FontMapEntry, tag: string | number) => {
      map[tag as string].xform.render(xmlStream, model![defn.prop]);
    });
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.map[node.name]) {
      this.parser = this.map[node.name].xform;
      return this.parser.parseOpen(node) as boolean;
    }
    switch (node.name) {
      case this.options.tagName:
        this.model = {};
        return true;
      default:
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser && !this.parser.parseClose(name)) {
      const item = this.map[name as string];
      if (this.parser.model) {
        (this.model as Record<string, unknown>)[item.prop] = this.parser.model;
      }
      this.parser = undefined;
      return true;
    }
    switch (name) {
      case this.options.tagName:
        return false;
      default:
        return true;
    }
  }

  static OPTIONS: FontXformOptions = {
    tagName: 'font',
    fontNameTag: 'name',
  };
}

export default FontXform;
export {FontXform};
