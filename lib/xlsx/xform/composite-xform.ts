import BaseXform from './base-xform.js';

/* 'virtual' methods used as a form of documentation */

// base class for xforms that are composed of other xforms
// offers some default implementations
class CompositeXform<TModel = Record<string, unknown>> extends BaseXform<TModel> {
  declare map: Record<string, BaseXform>;
}

export default CompositeXform;
export {CompositeXform};
