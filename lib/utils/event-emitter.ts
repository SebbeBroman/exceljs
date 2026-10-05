/**
 * Minimal EventEmitter compatible with the subset of Node's events API
 * used by this library (on/once/off/emit/removeAllListeners/addListener).
 * Replaces the npm `events` package so browser bundles stay free of that polyfill.
 *
 * Implemented as a function constructor (not `class`) so callers that use
 * `EventEmitter.call(this)` + `utils.inherits` (StreamBuf) keep working.
 */

export type EventListener = (...args: unknown[]) => void;

interface OnceWrapper extends EventListener {
  listener: EventListener;
}

type ListenerEntry = EventListener | OnceWrapper;

/** Instance shape of the EventEmitter constructor. */
export interface EventEmitterInstance {
  _events: Record<string, ListenerEntry | ListenerEntry[] | undefined>;
  on(type: string, listener: EventListener): this;
  addListener(type: string, listener: EventListener): this;
  once(type: string, listener: EventListener): this;
  off(type: string, listener: EventListener): this;
  removeListener(type: string, listener: EventListener): this;
  removeAllListeners(type?: string): this;
  emit(type: string, ...args: unknown[]): boolean;
  listenerCount(type: string): number;
  listeners(type: string): EventListener[];
}

/** @deprecated Use EventEmitterInstance */
export type EventEmitter = EventEmitterInstance;

export interface EventEmitterConstructor {
  new (): EventEmitterInstance;
  (): EventEmitterInstance;
  prototype: EventEmitterInstance;
}

function ensureListeners(
  ee: EventEmitter,
): Record<string, ListenerEntry | ListenerEntry[] | undefined> {
  if (!ee._events || typeof ee._events !== 'object') {
    ee._events = Object.create(null) as Record<string, ListenerEntry | ListenerEntry[] | undefined>;
  }
  return ee._events;
}

function arrayClone(arr: ListenerEntry[], n: number): ListenerEntry[] {
  const copy = Array.from({length: n}) as ListenerEntry[];
  for (let i = 0; i < n; i++) copy[i] = arr[i];
  return copy;
}

// Function constructor (not class) — required for EventEmitter.call(this) in StreamBuf
const EventEmitter = function EventEmitter(this: EventEmitter) {
  this._events = Object.create(null) as Record<string, ListenerEntry | ListenerEntry[] | undefined>;
} as EventEmitterConstructor;

EventEmitter.prototype.on = function on(this: EventEmitter, type: string, listener: EventListener) {
  return this.addListener(type, listener);
};

EventEmitter.prototype.addListener = function addListener(
  this: EventEmitter,
  type: string,
  listener: EventListener,
) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function');
  }
  const events = ensureListeners(this);
  const existing = events[type];
  if (!existing) {
    events[type] = listener;
  } else if (typeof existing === 'function') {
    events[type] = [existing, listener];
  } else {
    existing.push(listener);
  }
  return this;
};

EventEmitter.prototype.once = function once(
  this: EventEmitter,
  type: string,
  listener: EventListener,
) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function');
  }
  const ee = this;
  function wrapped(this: unknown, ...args: unknown[]) {
    ee.removeListener(type, wrapped);
    listener.apply(ee, args);
  }
  (wrapped as OnceWrapper).listener = listener;
  this.on(type, wrapped);
  return this;
};

EventEmitter.prototype.off = function off(
  this: EventEmitter,
  type: string,
  listener: EventListener,
) {
  return this.removeListener(type, listener);
};

EventEmitter.prototype.removeListener = function removeListener(
  this: EventEmitter,
  type: string,
  listener: EventListener,
) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function');
  }
  const events = this._events;
  if (!events) return this;
  const list = events[type];
  if (!list) return this;

  if (typeof list === 'function') {
    if (list === listener || (list as OnceWrapper).listener === listener) {
      delete events[type];
    }
    return this;
  }

  for (let i = 0; i < list.length; i++) {
    const item = list[i];
    if (item === listener || (item as OnceWrapper).listener === listener) {
      if (list.length === 1) {
        delete events[type];
      } else {
        list.splice(i, 1);
      }
      break;
    }
  }
  return this;
};

EventEmitter.prototype.removeAllListeners = function removeAllListeners(
  this: EventEmitter,
  type?: string,
) {
  const events = this._events;
  if (!events) return this;
  if (type === undefined) {
    this._events = Object.create(null) as Record<
      string,
      ListenerEntry | ListenerEntry[] | undefined
    >;
  } else {
    delete events[type];
  }
  return this;
};

EventEmitter.prototype.emit = function emit(
  this: EventEmitter,
  type: string,
  ...args: unknown[]
): boolean {
  const events = this._events;
  const handler = events && events[type];

  if (!handler) {
    if (type === 'error') {
      const err = args[0];
      if (err instanceof Error) throw err;
      const e = new Error('Unhandled error.') as Error & {context?: unknown};
      e.context = err;
      throw e;
    }
    return false;
  }

  if (typeof handler === 'function') {
    handler.apply(this, args);
  } else {
    const len = handler.length;
    const listeners = arrayClone(handler, len);
    for (let i = 0; i < len; i++) {
      listeners[i].apply(this, args);
    }
  }
  return true;
};

EventEmitter.prototype.listenerCount = function listenerCount(
  this: EventEmitter,
  type: string,
): number {
  const events = this._events;
  if (!events) return 0;
  const list = events[type];
  if (!list) return 0;
  if (typeof list === 'function') return 1;
  return list.length;
};

EventEmitter.prototype.listeners = function listeners(
  this: EventEmitter,
  type: string,
): EventListener[] {
  const events = this._events;
  if (!events) return [];
  const list = events[type];
  if (!list) return [];
  if (typeof list === 'function') {
    return [(list as OnceWrapper).listener || list];
  }
  return list.map(fn => (fn as OnceWrapper).listener || fn);
};

export default EventEmitter;
export {EventEmitter};
