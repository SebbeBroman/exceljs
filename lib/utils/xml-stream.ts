import _ from './under-dash.js';
import utils from './utils.js';

// constants
const OPEN_ANGLE = '<';
const CLOSE_ANGLE = '>';
const OPEN_ANGLE_SLASH = '</';
const CLOSE_SLASH_ANGLE = '/>';

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
function pushAttributes(xml: string[], attributes: XmlAttributes): void {
  if (attributes) {
    const tmp: string[] = [];
    _.each(attributes as Record<string, unknown>, (value, name) => {
      if (value !== undefined) {
        pushAttribute(tmp, name as string, value);
      }
    });
    xml.push(tmp.join(''));
  }
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

    // start streaming node
    xml.push(OPEN_ANGLE);
    xml.push(name);
    pushAttributes(xml, attributes);
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
      xml.push(OPEN_ANGLE_SLASH);
      xml.push(node!);
      xml.push(CLOSE_ANGLE);
    }
    this.open = false;
    this.leaf = false;
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: unknown): void {
    this.openNode(name, attributes);
    if (text !== undefined) {
      // zeros need to be written
      this.writeText(text);
    }
    this.closeNode();
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
