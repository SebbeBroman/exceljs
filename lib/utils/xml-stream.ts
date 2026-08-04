import utils from './utils.js';

// constants
const OPEN_ANGLE = '<';
const CLOSE_ANGLE = '>';
const OPEN_ANGLE_SLASH = '</';
const CLOSE_SLASH_ANGLE = '/>';

// Matches utils.xmlEncode special chars — used to skip encode on hot path
const XML_SPECIAL = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;

type XmlAttributes = Record<string, unknown> | null | undefined;

interface RollbackState {
  xml: number;
  stack: number;
  leaf: boolean;
  open: boolean;
}

function pushAttribute(xml: string[], name: string, value: unknown): void {
  xml.push(` ${name}="${utils.xmlEncode(String(value))}"`);
}

/** Build attribute string in one allocation (no per-attr array). */
function attributesToString(attributes: XmlAttributes): string {
  if (!attributes) return '';
  let out = '';
  for (const name of Object.keys(attributes)) {
    const value = attributes[name];
    if (value !== undefined) {
      out += ` ${name}="${utils.xmlEncode(String(value))}"`;
    }
  }
  return out;
}

function pushAttributes(xml: string[], attributes: XmlAttributes): void {
  const s = attributesToString(attributes);
  if (s) xml.push(s);
}

/** Encode leaf text; numbers and simple strings skip the full xmlEncode path. */
function encodeLeafText(text: unknown): string {
  if (typeof text === 'number') {
    // numbers never contain XML-special chars
    return String(text);
  }
  const s = String(text);
  // short plain strings (common for '0'/'1', shared-string ids) skip encode
  if (s.length <= 32 && !XML_SPECIAL.test(s)) {
    return s;
  }
  return utils.xmlEncode(s);
}

class XmlStream {
  private _xml: string[];
  private _stack: string[];
  private _rollbacks: RollbackState[];
  leaf: boolean;
  open: boolean;

  static StdDocAttributes: Record<string, string> = {
    version: '1.0',
    encoding: 'UTF-8',
    standalone: 'yes',
  };

  constructor() {
    this._xml = [];
    this._stack = [];
    this._rollbacks = [];
    this.leaf = false;
    this.open = false;
  }

  get tos(): string | undefined {
    return this._stack.length ? this._stack[this._stack.length - 1] : undefined;
  }

  get cursor(): number {
    // handy way to track whether anything has been added
    return this._xml.length;
  }

  openXml(docAttributes: XmlAttributes): void {
    const xml = this._xml;
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    xml.push('<?xml');
    pushAttributes(xml, docAttributes);
    xml.push('?>\n');
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    const parent = this.tos;
    const xml = this._xml;
    if (parent && this.open) {
      xml.push(CLOSE_ANGLE);
    }

    this._stack.push(name);

    // start streaming node — single push for name (+ attrs when present)
    const attrs = attributesToString(attributes);
    if (attrs) {
      xml.push(OPEN_ANGLE + name + attrs);
    } else {
      xml.push(OPEN_ANGLE + name);
    }
    this.leaf = true;
    this.open = true;
  }

  addAttribute(name: string, value: unknown): void {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    if (value !== undefined) {
      pushAttribute(this._xml, name, value);
    }
  }

  addAttributes(attrs: XmlAttributes): void {
    if (!this.open) {
      throw new Error('Cannot write attributes to node if it is not open');
    }
    pushAttributes(this._xml, attrs);
  }

  writeText(text: unknown): void {
    const xml = this._xml;
    if (this.open) {
      xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    xml.push(utils.xmlEncode(String(text)));
  }

  writeXml(xml: string): void {
    if (this.open) {
      this._xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;
    this._xml.push(xml);
  }

  closeNode(): void {
    const node = this._stack.pop();
    const xml = this._xml;
    if (this.leaf) {
      xml.push(CLOSE_SLASH_ANGLE);
    } else {
      // single chunk: </name>
      xml.push(OPEN_ANGLE_SLASH + node! + CLOSE_ANGLE);
    }
    this.open = false;
    this.leaf = false;
  }

  /**
   * Emit a complete leaf element in as few chunks as possible.
   * Hot path for sheet cells (`<v>…</v>`, self-closing tags, etc.).
   * Does not touch the element stack (complete element written atomically).
   * Rollback/commit semantics preserved via _xml length only.
   */
  leafNode(name: string, attributes?: XmlAttributes, text?: unknown): void {
    const xml = this._xml;
    // Close any currently open start-tag before writing sibling content
    if (this.open) {
      xml.push(CLOSE_ANGLE);
      this.open = false;
    }
    this.leaf = false;

    const attrs = attributesToString(attributes);
    if (text === undefined) {
      // zeros need to be written — but undefined means self-closing
      xml.push(OPEN_ANGLE + name + attrs + CLOSE_SLASH_ANGLE);
      return;
    }

    // Single chunk: <name attrs>text</name>
    // numbers / short plain strings skip full xmlEncode
    xml.push(OPEN_ANGLE + name + attrs + CLOSE_ANGLE + encodeLeafText(text) + OPEN_ANGLE_SLASH + name + CLOSE_ANGLE);
  }

  closeAll(): void {
    while (this._stack.length) {
      this.closeNode();
    }
  }

  addRollback(): number {
    this._rollbacks.push({
      xml: this._xml.length,
      stack: this._stack.length,
      leaf: this.leaf,
      open: this.open,
    });
    return this.cursor;
  }

  commit(): void {
    this._rollbacks.pop();
  }

  rollback(): void {
    const r = this._rollbacks.pop()!;
    if (this._xml.length > r.xml) {
      this._xml.splice(r.xml, this._xml.length - r.xml);
    }
    if (this._stack.length > r.stack) {
      this._stack.splice(r.stack, this._stack.length - r.stack);
    }
    this.leaf = r.leaf;
    this.open = r.open;
  }

  get xml(): string {
    this.closeAll();
    return this._xml.join('');
  }
}

export default XmlStream;
export {XmlStream};
