import type {DateXformOptions} from '../../xform/simple/date-xform.js';
export type {DateXformOptions} from '../../xform/simple/date-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class DateXform extends BaseXform<Date> {
  attr: string | undefined;
  attrs: Record<string, unknown> | undefined;
  _format: (dt: Date) => string;
  _parse: (str: string) => Date;
  text?: string[];

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
    this._parse =
      options.parse ||
      function (str: string) {
        return new Date(str);
      };
  }

  override parseOpen(node: XmlNode): void {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = this._parse(node.attributes[this.attr]);
      } else {
        this.text = [];
      }
    }
  }

  override parseText(text: string): void {
    if (!this.attr) {
      this.text!.push(text);
    }
  }

  override parseClose(): boolean {
    if (!this.attr) {
      this.model = this._parse(this.text!.join(''));
    }
    return false;
  }
}

export default DateXform;
export {DateXform};
