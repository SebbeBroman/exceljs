import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface DateXformOptions {
  tag: string;
  attr?: string;
  attrs?: Record<string, unknown>;
  format?: (dt: Date) => string;
  parse?: (str: string) => Date;
}

class DateXform extends BaseXform<Date> {
  attr: string | undefined;
  attrs: Record<string, unknown> | undefined;
  _format: (dt: Date) => string;

  constructor(options: DateXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.attrs = options.attrs;
    this._format =
      options.format ||
      function (dt: Date) {
        try {
          if (Number.isNaN(dt.getTime())) return '';
          return dt.toISOString();
        } catch (_e) {
          return '';
        }
      };
  }

  override render(xmlStream: XmlStreamLike, model?: Date | null): void {
    if (model) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, this._format(model));
      } else {
        xmlStream.writeText(this._format(model));
      }
      xmlStream.closeNode();
    }
  }
}

export default DateXform;
export {DateXform};
