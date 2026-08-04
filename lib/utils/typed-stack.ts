// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Constructor<T> = new (...args: any[]) => T;

class TypedStack<T> {
  private _type: Constructor<T>;
  private _stack: T[];

  constructor(type: Constructor<T>) {
    this._type = type;
    this._stack = [];
  }

  get size(): number {
    return this._stack.length;
  }

  pop(): T {
    const tos = this._stack.pop();
    return tos || new this._type();
  }

  push(instance: T): void {
    if (!(instance instanceof this._type)) {
      throw new Error('Invalid type pushed to TypedStack');
    }
    this._stack.push(instance);
  }
}

export default TypedStack;
export {TypedStack};
