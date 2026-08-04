/**
 * Minimal Chai BDD-style expect() built on Vitest's expect.
 * Supports dirty-chai style both as property and as function:
 *   expect(x).to.be.ok
 *   expect(x).to.be.ok()
 *   expect(x).to.not.be.undefined()
 */

function typeName(value) {
  if (value === null) return 'null';
  if (Array.isArray(value)) return 'array';
  if (value instanceof Error) return 'error';
  return typeof value;
}

function normalizeXml(xml) {
  let s = String(xml)
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .replace(/>\s+/g, '>')
    .replace(/\s+</g, '<')
    .replace(/\s+\/>/g, '/>') // <tag /> → <tag/>
    .trim();
  // Sort attributes so order differences don't fail (chai-xml was order-tolerant)
  s = s.replace(/<([A-Za-z0-9:_-]+)((?:\s+[A-Za-z0-9:_-]+="[^"]*")*)(\s*\/?)>/g, (_m, tag, attrs, close) => {
    const selfClose = close.includes('/') ? '/' : '';
    if (!attrs || !attrs.trim()) return `<${tag}${selfClose}>`;
    const list = [];
    const re = /([A-Za-z0-9:_-]+)="([^"]*)"/g;
    let m;
    while ((m = re.exec(attrs))) {
      list.push(`${m[1]}="${m[2]}"`);
    }
    list.sort();
    return `<${tag}${list.length ? ` ${list.join(' ')}` : ''}${selfClose}>`;
  });
  // <tag attrs></tag> with no content → <tag attrs/>
  s = s.replace(/<([A-Za-z0-9:_-]+)((?:\s+[A-Za-z0-9:_-]+="[^"]*")*)><\/\1>/g, '<$1$2/>');
  return s;
}

/** Allow dirty-chai: property assert that may also be invoked as a function. */
function chainable(assertion, run) {
  run();
  const fn = () => assertion;
  return new Proxy(fn, {
    get(_t, prop) {
      if (prop === 'then') return undefined; // not thenable
      if (prop === Symbol.toStringTag) return 'Assertion';
      const val = assertion[prop];
      return typeof val === 'function' ? val.bind(assertion) : val;
    },
  });
}

export function createChaiExpect(viExpect) {
  function expect(actual, message) {
    return new Assertion(actual, {message}, viExpect);
  }
  expect.fail = msg => {
    throw new Error(msg || 'expect.fail');
  };
  return expect;
}

class Assertion {
  constructor(actual, flags, viExpect) {
    this._actual = actual;
    this._flags = flags || {};
    this._vi = viExpect;
  }

  get to() {
    return this;
  }
  get be() {
    return this;
  }
  get been() {
    return this;
  }
  get is() {
    return this;
  }
  get that() {
    return this;
  }
  get and() {
    return this;
  }
  get have() {
    return this;
  }
  get has() {
    return this;
  }
  get with() {
    return this;
  }
  get at() {
    return this;
  }
  get of() {
    return this;
  }
  get same() {
    return this;
  }

  get not() {
    return new Assertion(this._actual, {...this._flags, not: !this._flags.not}, this._vi);
  }

  get deep() {
    return new Assertion(this._actual, {...this._flags, deep: true}, this._vi);
  }

  get xml() {
    return new Assertion(this._actual, {...this._flags, xml: true}, this._vi);
  }

  get ok() {
    return chainable(this, () => {
      if (this._flags.not) this._vi(this._actual).toBeFalsy();
      else this._vi(this._actual).toBeTruthy();
    });
  }

  get true() {
    return chainable(this, () => {
      if (this._flags.not) this._vi(this._actual).not.toBe(true);
      else this._vi(this._actual).toBe(true);
    });
  }

  get false() {
    return chainable(this, () => {
      if (this._flags.not) this._vi(this._actual).not.toBe(false);
      else this._vi(this._actual).toBe(false);
    });
  }

  get null() {
    return chainable(this, () => {
      if (this._flags.not) this._vi(this._actual).not.toBeNull();
      else this._vi(this._actual).toBeNull();
    });
  }

  get undefined() {
    return chainable(this, () => {
      if (this._flags.not) this._vi(this._actual).not.toBeUndefined();
      else this._vi(this._actual).toBeUndefined();
    });
  }

  get exist() {
    return chainable(this, () => {
      const exists = this._actual !== null && this._actual !== undefined;
      if (this._flags.not) this._vi(exists).toBe(false);
      else this._vi(exists).toBe(true);
    });
  }

  get empty() {
    return chainable(this, () => {
      const len =
        this._actual == null
          ? 0
          : typeof this._actual === 'string' || Array.isArray(this._actual)
            ? this._actual.length
            : Object.keys(this._actual).length;
      if (this._flags.not) this._vi(len).not.toBe(0);
      else this._vi(len).toBe(0);
    });
  }

  equal(expected) {
    if (this._flags.xml) {
      const a = normalizeXml(this._actual);
      const b = normalizeXml(expected);
      if (this._flags.not) this._vi(a).not.toBe(b);
      else this._vi(a).toBe(b);
      return this;
    }
    if (this._flags.deep) {
      if (this._flags.not) this._vi(this._actual).not.toEqual(expected);
      else this._vi(this._actual).toEqual(expected);
    } else if (this._flags.not) {
      this._vi(this._actual).not.toBe(expected);
    } else {
      this._vi(this._actual).toBe(expected);
    }
    return this;
  }

  eql(expected) {
    return this.deep.equal(expected);
  }
  equals(expected) {
    return this.equal(expected);
  }
  eq(expected) {
    return this.equal(expected);
  }

  equalDate(expected) {
    const a = this._actual instanceof Date ? this._actual.getTime() : new Date(this._actual).getTime();
    const b = expected instanceof Date ? expected.getTime() : new Date(expected).getTime();
    if (this._flags.not) this._vi(a).not.toBe(b);
    else this._vi(a).toBe(b);
    return this;
  }

  match(re) {
    if (this._flags.not) this._vi(this._actual).not.toMatch(re);
    else this._vi(this._actual).toMatch(re);
    return this;
  }

  throw(expected) {
    const fn = this._actual;
    if (typeof fn !== 'function') {
      throw new Error('expected a function to test for throw');
    }
    if (this._flags.not) this._vi(fn).not.toThrow(expected);
    else if (expected === undefined) this._vi(fn).toThrow();
    else this._vi(fn).toThrow(expected);
    return this;
  }
  throws(expected) {
    return this.throw(expected);
  }

  property(name, value) {
    if (this._flags.not) this._vi(this._actual).not.toHaveProperty(name);
    else if (arguments.length > 1) this._vi(this._actual).toHaveProperty(name, value);
    else this._vi(this._actual).toHaveProperty(name);
    return this;
  }

  length(n) {
    if (this._flags.not) this._vi(this._actual).not.toHaveLength(n);
    else this._vi(this._actual).toHaveLength(n);
    return this;
  }
  lengthOf(n) {
    return this.length(n);
  }

  members(expected) {
    const actual = this._actual;
    if (this._flags.not) {
      const hasAll = expected.every(e => actual.includes(e));
      this._vi(hasAll).toBe(false);
    } else {
      for (const e of expected) this._vi(actual).toContain(e);
    }
    return this;
  }

  below(n) {
    if (this._flags.not) this._vi(this._actual).not.toBeLessThan(n);
    else this._vi(this._actual).toBeLessThan(n);
    return this;
  }
  lessThan(n) {
    return this.below(n);
  }
  greaterThan(n) {
    if (this._flags.not) this._vi(this._actual).not.toBeGreaterThan(n);
    else this._vi(this._actual).toBeGreaterThan(n);
    return this;
  }
  above(n) {
    return this.greaterThan(n);
  }

  a(type) {
    const actualType = typeName(this._actual);
    const expected = String(type).toLowerCase();
    if (expected === 'array') {
      if (this._flags.not) this._vi(Array.isArray(this._actual)).toBe(false);
      else this._vi(Array.isArray(this._actual)).toBe(true);
    } else if (this._flags.not) this._vi(actualType).not.toBe(expected);
    else this._vi(actualType).toBe(expected);
    return this;
  }
  an(type) {
    return this.a(type);
  }

  get valid() {
    return chainable(this, () => {
      if (!this._flags.xml) {
        throw new Error('.valid() is only supported after .xml');
      }
      const s = String(this._actual);
      this._vi(s.includes('<') && s.includes('>')).toBe(true);
    });
  }
}
