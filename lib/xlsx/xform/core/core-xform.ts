import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import DateXform from '../simple/date-xform.js';
import StringXform from '../simple/string-xform.js';
import IntegerXform from '../simple/integer-xform.js';

export interface CoreModel {
  creator?: string;
  title?: string;
  subject?: string;
  description?: string;
  identifier?: string;
  language?: string;
  keywords?: string;
  category?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  revision?: number;
  version?: string;
  contentStatus?: string;
  contentType?: string;
  created?: Date;
  modified?: Date;
}

class CoreXform extends BaseXform<CoreModel> {
  declare map: Record<string, BaseXform>;

  constructor() {
    super();

    this.map = {
      'dc:creator': new StringXform({tag: 'dc:creator'}),
      'dc:title': new StringXform({tag: 'dc:title'}),
      'dc:subject': new StringXform({tag: 'dc:subject'}),
      'dc:description': new StringXform({tag: 'dc:description'}),
      'dc:identifier': new StringXform({tag: 'dc:identifier'}),
      'dc:language': new StringXform({tag: 'dc:language'}),
      'cp:keywords': new StringXform({tag: 'cp:keywords'}),
      'cp:category': new StringXform({tag: 'cp:category'}),
      'cp:lastModifiedBy': new StringXform({tag: 'cp:lastModifiedBy'}),
      'cp:lastPrinted': new DateXform({tag: 'cp:lastPrinted', format: CoreXform.DateFormat}),
      'cp:revision': new IntegerXform({tag: 'cp:revision'}),
      'cp:version': new StringXform({tag: 'cp:version'}),
      'cp:contentStatus': new StringXform({tag: 'cp:contentStatus'}),
      'cp:contentType': new StringXform({tag: 'cp:contentType'}),
      'dcterms:created': new DateXform({
        tag: 'dcterms:created',
        attrs: CoreXform.DateAttrs,
        format: CoreXform.DateFormat,
      }),
      'dcterms:modified': new DateXform({
        tag: 'dcterms:modified',
        attrs: CoreXform.DateAttrs,
        format: CoreXform.DateFormat,
      }),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: CoreModel | null): void {
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);

    xmlStream.openNode('cp:coreProperties', CoreXform.CORE_PROPERTY_ATTRIBUTES);

    this.map['dc:creator'].render(xmlStream, model!.creator);
    this.map['dc:title'].render(xmlStream, model!.title);
    this.map['dc:subject'].render(xmlStream, model!.subject);
    this.map['dc:description'].render(xmlStream, model!.description);
    this.map['dc:identifier'].render(xmlStream, model!.identifier);
    this.map['dc:language'].render(xmlStream, model!.language);
    this.map['cp:keywords'].render(xmlStream, model!.keywords);
    this.map['cp:category'].render(xmlStream, model!.category);
    this.map['cp:lastModifiedBy'].render(xmlStream, model!.lastModifiedBy);
    this.map['cp:lastPrinted'].render(xmlStream, model!.lastPrinted);
    this.map['cp:revision'].render(xmlStream, model!.revision);
    this.map['cp:version'].render(xmlStream, model!.version);
    this.map['cp:contentStatus'].render(xmlStream, model!.contentStatus);
    this.map['cp:contentType'].render(xmlStream, model!.contentType);
    this.map['dcterms:created'].render(xmlStream, model!.created);
    this.map['dcterms:modified'].render(xmlStream, model!.modified);

    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'cp:coreProperties':
      case 'coreProperties':
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
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
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'cp:coreProperties':
      case 'coreProperties':
        this.model = {
          creator: this.map['dc:creator'].model as string | undefined,
          title: this.map['dc:title'].model as string | undefined,
          subject: this.map['dc:subject'].model as string | undefined,
          description: this.map['dc:description'].model as string | undefined,
          identifier: this.map['dc:identifier'].model as string | undefined,
          language: this.map['dc:language'].model as string | undefined,
          keywords: this.map['cp:keywords'].model as string | undefined,
          category: this.map['cp:category'].model as string | undefined,
          lastModifiedBy: this.map['cp:lastModifiedBy'].model as string | undefined,
          lastPrinted: this.map['cp:lastPrinted'].model as Date | undefined,
          revision: this.map['cp:revision'].model as number | undefined,
          contentStatus: this.map['cp:contentStatus'].model as string | undefined,
          contentType: this.map['cp:contentType'].model as string | undefined,
          created: this.map['dcterms:created'].model as Date | undefined,
          modified: this.map['dcterms:modified'].model as Date | undefined,
        };
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }

  static DateFormat(dt: Date): string {
    return dt.toISOString().replace(/[.]\d{3}/, '');
  }

  static DateAttrs = {'xsi:type': 'dcterms:W3CDTF'};

  static CORE_PROPERTY_ATTRIBUTES: Record<string, string> = {
    'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
    'xmlns:dcterms': 'http://purl.org/dc/terms/',
    'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
    'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
  };
}

export default CoreXform;
export {CoreXform};
