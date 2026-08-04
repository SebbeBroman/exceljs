import XmlStream from '../../utils/xml-stream.js';

/* 'virtual' methods used as a form of documentation */

// Lazy SAX: write-only paths never touch XML parsing. Dynamic import lets
// code-splitting bundlers keep saxen out of the initial chunk.
let parseSaxPromise: Promise<typeof import('../../utils/parse-sax.js').default> | undefined;
function loadParseSax(): Promise<typeof import('../../utils/parse-sax.js').default> {
  if (!parseSaxPromise) {
    parseSaxPromise = import('../../utils/parse-sax.js').then(m => m.default);
  }
  return parseSaxPromise;
}

/** Minimal XML stream interface used by xforms (matches utils/xml-stream). */
export interface XmlStreamLike {
  openXml(docAttributes?: Record<string, string>): void;
  openNode(name: string, attributes?: Record<string, unknown>): void;
  addAttribute(name: string, value: unknown): void;
  addAttributes(attrs: Record<string, unknown> | undefined | null): void;
  writeText(text: unknown): void;
  writeXml(xml: string): void;
  closeNode(): void;
  leafNode(name: string, attributes?: Record<string, unknown>, text?: unknown): void;
  closeAll(): void;
  addRollback(): number;
  commit(): void;
  rollback(): void;
  readonly xml: string;
  readonly cursor: number;
  readonly tos: string | undefined;
}

/** SAX open-tag node shape from parse-sax. */
export interface XmlNode {
  name: string;
  attributes: Record<string, string>;
}

export interface SaxEvent {
  eventType: 'opentag' | 'text' | 'closetag' | string;
  value: XmlNode | string | {name: string};
}

export type SaxParser = AsyncIterable<SaxEvent[]>;

export interface XformOptions {
  index?: number;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

// Base class for Xforms
// Note: map/parser/tag intentionally loosely typed so subclasses with getters
// and heterogeneous child maps type-check under strict TS.
class BaseXform<TModel = unknown> {
  // Backing field for `model` accessors. Named `_bxModel` (not `_model`) because
  // StaticXform / Vml* xforms historically store constructor options on `_model`.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected _bxModel: any;

  get model(): TModel | null | undefined {
    return this._bxModel as TModel | null | undefined;
  }

  set model(value: TModel | null | undefined) {
    this._bxModel = value;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  map?: Record<string, any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  parser?: any;

  /** XML element name; subclasses set a string field (not a getter). */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  tag?: any;

  // ============================================================
  // Virtual Interface
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  prepare(_model?: any, _options?: XformOptions): void {
    // optional preparation (mutation) of model so it is ready for write
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  render(_xmlStream: XmlStreamLike, _model?: any, _arg?: any): void {
    // convert model to xml
  }

  parseOpen(_node: XmlNode): boolean | void {
    // XML node opened
  }

  parseText(_text: string): void {
    // chunk of text encountered for current node
  }

  parseClose(_name?: string): boolean | void {
    // XML node closed
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  reconcile(_model?: any, _options?: any): void {
    // optional post-parse step (opposite to prepare)
  }

  // ============================================================
  reset(): void {
    // to make sure parses don't bleed to next iteration
    this.model = null;

    // if we have a map - reset them too
    if (this.map) {
      Object.values(this.map).forEach(xform => {
        if (xform instanceof BaseXform) {
          xform.reset();
        } else if (xform && typeof xform === 'object' && 'xform' in xform && xform.xform) {
          (xform.xform as BaseXform).reset();
        }
      });
    }
  }

  mergeModel(obj: Record<string, unknown>): void {
    // set obj's props to this.model
    this.model = Object.assign((this.model as object) || {}, obj) as TModel;
  }

  async parse(saxParser: SaxParser): Promise<TModel | null | undefined> {
    for await (const events of saxParser) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          this.parseOpen(value as XmlNode);
        } else if (eventType === 'text') {
          this.parseText(value as string);
        } else if (eventType === 'closetag') {
          if (!this.parseClose((value as {name: string}).name)) {
            return this.model;
          }
        }
      }
    }
    return this.model;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async parseStream(stream: any): Promise<TModel | null | undefined> {
    const parseSax = await loadParseSax();
    return this.parse(parseSax(stream) as SaxParser);
  }

  get xml(): string {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  toXml(model?: any): string {
    const xmlStream = new XmlStream() as unknown as XmlStreamLike;
    this.render(xmlStream, model);
    return xmlStream.xml;
  }

  // ============================================================
  // Useful Utilities
  static toAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt !== undefined && dflt !== null ? String(dflt) : (dflt as undefined);
      }
    } else if (always || value !== dflt) {
      return (value as {toString(): string}).toString();
    }
    return undefined;
  }

  static toStringAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toStringValue(attr: string | undefined, dflt?: string): string | undefined {
    return attr === undefined ? dflt : attr;
  }

  static toBoolAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt as string | undefined;
      }
    } else if (always || value !== dflt) {
      return value ? '1' : '0';
    }
    return undefined;
  }

  static toBoolValue(attr: string | undefined, dflt?: boolean): boolean | undefined {
    return attr === undefined ? dflt : attr === '1';
  }

  static toIntAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toIntValue(attr: string | undefined, dflt?: number): number | undefined {
    return attr === undefined ? dflt : parseInt(attr, 10);
  }

  static toFloatAttribute(value: unknown, dflt?: unknown, always = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toFloatValue(attr: string | undefined, dflt?: number): number | undefined {
    return attr === undefined ? dflt : parseFloat(attr);
  }
}

export default BaseXform;
export {BaseXform};
