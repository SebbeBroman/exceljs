
const tools = {
  dtMatcher: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[.]\d{3}Z$/,
  fix: function fix(o) {
    // clone the object and replace any date-like strings with new Date()
    let clone;
    if (o instanceof Array) {
      clone = [];
    } else if (typeof o === 'object') {
      clone = {};
    } else if (typeof o === 'string' && tools.dtMatcher.test(o)) {
      return new Date(o);
    } else {
      return o;
    }
    if (Array.isArray(o)) {
      o.forEach((value, name) => {
        if (value !== undefined) {
          clone[name] = fix(value);
        }
      });
    } else {
      for (const [name, value] of Object.entries(o)) {
        if (value !== undefined) {
          clone[name] = fix(value);
        }
      }
    }
    return clone;
  },

  concatenateFormula(...args) {
    const values = args.map(value => `"${value}"`);
    return {
      formula: `CONCATENATE(${values.join(',')})`,
    };
  },
};

export default tools;
