import PageBreaksXform from './page-breaks-xform.js';
import ListXform from '../list-xform.js';

class RowBreaksXform extends ListXform {
  constructor() {
    const options = {
      tag: 'rowBreaks',
      count: true,
      childXform: new PageBreaksXform(),
    };
    super(options);
  }
}

export default RowBreaksXform;
export {RowBreaksXform};
