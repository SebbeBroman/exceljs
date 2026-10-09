import type {ExtLstModel} from '../../xform/sheet/ext-lst-xform.js';
export type {ExtLstModel} from '../../xform/sheet/ext-lst-xform.js';
import CompositeXform from '../composite-xform.js';
import ConditionalFormattingsExt from './cf-ext/conditional-formattings-ext-xform.js';
import type {ConditionalFormattingsExtModel} from './cf-ext/conditional-formattings-ext-xform.js';

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

  hasContent(model?: ExtLstModel | null): boolean {
    return this.ext.hasContent(model);
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
