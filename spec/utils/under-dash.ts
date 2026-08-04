// @vitest-migrated
import libUnderDash from '../../lib/utils/under-dash.js';

const _ = Object.assign(
  {
    get: function get(obj, path, dflt) {
      if (typeof path === 'string') {
        path = path.split('.');
      }
      while (obj && path.length) {
        obj = obj[path.shift()];
      }
      return obj !== undefined ? obj : dflt;
    },

    has: function has(obj, path) {
      const dummy = {};
      return _.get(obj, path, dummy) !== dummy;
    },

    cloneDeep: function cloneDeep(obj, preserveUndefined) {
      if (preserveUndefined === undefined) {
        preserveUndefined = true;
      }
      let clone;
      if (obj === null) {
        return null;
      }
      if (obj instanceof Date) {
        return obj;
      }
      if (obj instanceof Array) {
        clone = [];
      } else if (typeof obj === 'object') {
        clone = {};
      } else {
        return obj;
      }
      _.each(obj, (value, name) => {
        if (value !== undefined) {
          clone[name] = cloneDeep(value, preserveUndefined);
        } else if (preserveUndefined) {
          clone[name] = undefined;
        }
      });
      return clone;
    },
  },
  libUnderDash
);

export default _;
// Named re-exports for `import {each, cloneDeep} from ...` (migrated CJS destructure)
export const each = (...args) => _.each(...args);
export const get = (...args) => _.get(...args);
export const has = (...args) => _.has(...args);
export const cloneDeep = (...args) => _.cloneDeep(...args);
export const isEqual = (...args) => _.isEqual(...args);
export const isObject = (...args) => _.isObject(...args);
export const isUndefined = (...args) => _.isUndefined(...args);
export const some = (...args) => _.some(...args);
export const every = (...args) => _.every(...args);
export const map = (...args) => _.map(...args);
export const keyBy = (...args) => _.keyBy(...args);
export const strcmp = (...args) => _.strcmp(...args);
export const escapeHtml = (...args) => _.escapeHtml(...args);
export const deepMerge = (...args) => _.deepMerge(...args);
