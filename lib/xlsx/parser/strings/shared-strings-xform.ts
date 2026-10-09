import type {SharedStringsModel} from '../../xform/strings/shared-strings-xform.js';
export type {SharedStringsModel} from '../../xform/strings/shared-strings-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import SharedStringXform from './shared-string-xform.js';
import type {SharedStringModel} from './shared-string-xform.js';

class SharedStringsXform extends BaseXform<SharedStringsModel> {
  _sharedStringXform?: SharedStringXform;

  constructor(model?: SharedStringsModel) {
    super();

    this.model = model || {
      values: [],
      count: 0,
    };
  }

  get sharedStringXform(): SharedStringXform {
    return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
  }

  get values(): SharedStringModel[] {
    return this.model!.values;
  }

  get uniqueCount(): number {
    return this.model!.values.length;
  }

  get count(): number {
    return this.model!.count;
  }

  getString(index: number): SharedStringModel {
    return this.model!.values[index];
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'sst':
        return true;
      case 'si':
        this.parser = this.sharedStringXform;
        this.parser.parseOpen(node);
        return true;
      default:
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.values.push(this.parser.model as SharedStringModel);
        this.model!.count++;
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'sst':
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export default SharedStringsXform;
export {SharedStringsXform};
