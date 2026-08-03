import BaseXform from '../base-xform.js';

class PageBreaksXform extends BaseXform {
  get tag() {
    return 'brk';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('brk', model);
  }

  parseOpen(node) {
    if (node.name === 'brk') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default PageBreaksXform;
export { PageBreaksXform };
