import type {StaticXmlModel} from '../xform/static-xform.js';
export type {StaticXmlModel} from '../xform/static-xform.js';
import BaseXform from '../base-parser.js';

/** Consume an ignored XML subtree without retaining its rendering template. */
class StaticXform extends BaseXform {
  private root: string;
  constructor({tag}: Pick<StaticXmlModel, 'tag'>) {
    super();
    this.root = tag;
  }
  override parseOpen(): boolean {
    return true;
  }
  override parseText(): void {}
  override parseClose(name?: string): boolean {
    return name !== this.root;
  }
}
export default StaticXform;
export {StaticXform};
