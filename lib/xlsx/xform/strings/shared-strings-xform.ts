import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import SharedStringXform from './shared-string-xform.js';
import type {SharedStringModel} from './shared-string-xform.js';

export interface SharedStringsModel {
  values: SharedStringModel[];
  count: number;
}

class SharedStringsXform extends BaseXform<SharedStringsModel> {
  hash: Record<string, number>;
  rich: Record<string, number>;
  _sharedStringXform?: SharedStringXform;

  constructor(model?: SharedStringsModel) {
    super();

    this.model = model || {
      values: [],
      count: 0,
    };
    this.hash = Object.create(null) as Record<string, number>;
    this.rich = Object.create(null) as Record<string, number>;
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

  add(value: SharedStringModel): number {
    return value && typeof value === 'object' && 'richText' in value
      ? this.addRichText(value)
      : this.addText(value as string);
  }

  addText(value: string): number {
    let index = this.hash[value];
    if (index === undefined) {
      index = this.hash[value] = this.model!.values.length;
      this.model!.values.push(value);
    }
    this.model!.count++;
    return index;
  }

  addRichText(value: SharedStringModel): number {
    // TODO: add WeakMap here
    const xml = this.sharedStringXform.toXml(value);
    let index = this.rich[xml];
    if (index === undefined) {
      index = this.rich[xml] = this.model!.values.length;
      this.model!.values.push(value);
    }
    this.model!.count++;
    return index;
  }

  // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
  //   <si><t><%=text%></t></si>
  //   <si><r><rPr></rPr><t></t></r></si>
  // </sst>

  override render(xmlStream: XmlStreamLike, model?: SharedStringsModel | null): void {
    model = model || this.model!;
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);

    xmlStream.openNode('sst', {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      count: model.count,
      uniqueCount: model.values.length,
    });

    const sx = this.sharedStringXform;
    model.values.forEach(sharedString => {
      sx.render(xmlStream, sharedString);
    });
    xmlStream.closeNode();
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
