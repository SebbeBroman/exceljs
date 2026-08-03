/**
 * Minimal EventEmitter compatible with the subset of Node's events API
 * used by this library (on/once/off/emit/removeAllListeners/addListener).
 * Replaces the npm `events` package so browser bundles stay free of that polyfill.
 *
 * Implemented as a function constructor (not `class`) so callers that use
 * `EventEmitter.call(this)` + `utils.inherits` (StreamBuf) keep working.
 */

function ensureListeners(ee) {
  if (!ee._events || typeof ee._events !== 'object') {
    ee._events = Object.create(null);
  }
  return ee._events;
}

function arrayClone(arr, n) {
  const copy = new Array(n);
  for (let i = 0; i < n; i++) copy[i] = arr[i];
  return copy;
}

function EventEmitter() {
  this._events = Object.create(null);
}

EventEmitter.prototype.on = function on(type, listener) {
  return this.addListener(type, listener);
};

EventEmitter.prototype.addListener = function addListener(type, listener) {
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

EventEmitter.prototype.once = function once(type, listener) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function');
  }
  const ee = this;
  function wrapped(...args) {
    ee.removeListener(type, wrapped);
    listener.apply(ee, args);
  }
  wrapped.listener = listener;
  this.on(type, wrapped);
  return this;
};

EventEmitter.prototype.off = function off(type, listener) {
  return this.removeListener(type, listener);
};

EventEmitter.prototype.removeListener = function removeListener(type, listener) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function');
  }
  const events = this._events;
  if (!events) return this;
  const list = events[type];
  if (!list) return this;

  if (typeof list === 'function') {
    if (list === listener || list.listener === listener) {
      delete events[type];
    }
    return this;
  }

  for (let i = 0; i < list.length; i++) {
    const item = list[i];
    if (item === listener || item.listener === listener) {
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

EventEmitter.prototype.removeAllListeners = function removeAllListeners(type) {
  const events = this._events;
  if (!events) return this;
  if (type === undefined) {
    this._events = Object.create(null);
  } else {
    delete events[type];
  }
  return this;
};

EventEmitter.prototype.emit = function emit(type, ...args) {
  const events = this._events;
  const handler = events && events[type];

  if (!handler) {
    if (type === 'error') {
      const err = args[0];
      if (err instanceof Error) throw err;
      const e = new Error('Unhandled error.');
      // eslint-disable-next-line no-underscore-dangle
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

EventEmitter.prototype.listenerCount = function listenerCount(type) {
  const events = this._events;
  if (!events) return 0;
  const list = events[type];
  if (!list) return 0;
  if (typeof list === 'function') return 1;
  return list.length;
};

EventEmitter.prototype.listeners = function listeners(type) {
  const events = this._events;
  if (!events) return [];
  const list = events[type];
  if (!list) return [];
  if (typeof list === 'function') {
    return [list.listener || list];
  }
  return list.map(fn => fn.listener || fn);
};

export default EventEmitter;
export {EventEmitter};
