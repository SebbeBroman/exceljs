import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class AppHeadingPairsXform extends BaseXform<unknown[]> {
  override parseOpen(node: XmlNode): boolean {
    // no parsing
    return node.name === 'HeadingPairs';
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    return name !== 'HeadingPairs';
  }
}

export default AppHeadingPairsXform;
export {AppHeadingPairsXform};
