import CompositeXform from '../composite-xform.js';
import type {XmlStreamLike, XformOptions} from '../base-xform.js';
import ConditionalFormattingsExt from './cf-ext/conditional-formattings-ext-xform.js';
import type {ConditionalFormattingsExtModel} from './cf-ext/conditional-formattings-ext-xform.js';

export interface ExtLstModel {
  conditionalFormattings?: ConditionalFormattingsExtModel;
  [key: string]: unknown;
}

class ExtXform extends CompositeXform<ExtLstModel> {
  conditionalFormattings: ConditionalFormattingsExt;

  constructor() {
    super();
    this.map = {
      'x14:conditionalFormattings': (this.conditionalFormattings = new ConditionalFormattingsExt()),
    };
  }

  override tag = 'ext';

  hasContent(model?: ExtLstModel | null): boolean {
    return !!this.conditionalFormattings.hasContent(
      model?.conditionalFormattings as ConditionalFormattingsExtModel,
    );
  }

  override prepare(model?: ExtLstModel | null, options?: XformOptions): void {
    if (!model) {
      return;
    }
    this.conditionalFormattings.prepare(model.conditionalFormattings, options);
  }

  override render(xmlStream: XmlStreamLike, model?: ExtLstModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode('ext', {
      uri: '{78C0D931-6437-407d-A8EE-F0AAD7539E65}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });

    this.conditionalFormattings.render(xmlStream, model.conditionalFormattings);

    xmlStream.closeNode();
  }

  override createNewModel(): ExtLstModel {
    return {};
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    (this.model as ExtLstModel)[name] = parser.model;
  }
}

class ExtLstXform extends CompositeXform<ExtLstModel> {
  ext: ExtXform;

  constructor() {
    super();

    this.map = {
      ext: (this.ext = new ExtXform()),
    };
  }

  override tag = 'extLst';

  override prepare(model?: ExtLstModel | null, options?: XformOptions): void {
    this.ext.prepare(model, options);
  }

  hasContent(model?: ExtLstModel | null): boolean {
    return this.ext.hasContent(model);
  }

  override render(xmlStream: XmlStreamLike, model?: ExtLstModel | null): void {
    if (!this.hasContent(model)) {
      return;
    }

    xmlStream.openNode('extLst');
    this.ext.render(xmlStream, model);
    xmlStream.closeNode();
  }

  override createNewModel(): ExtLstModel {
    return {};
  }

  override onParserClose(_name: string, parser: {model: unknown}): void {
    Object.assign(this.model as ExtLstModel, parser.model);
  }
}

export default ExtLstXform;
export {ExtLstXform};
