import type {CfRuleModel} from '../../../xform/sheet/cf/cf-rule-xform.js';
export type {CfRuleModel} from '../../../xform/sheet/cf/cf-rule-xform.js';
import BaseXform from '../../../base-parser.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlNode} from '../../../base-parser.js';
import DatabarXform from './databar-xform.js';
import ExtLstRefXform from './ext-lst-ref-xform.js';
import FormulaXform from './formula-xform.js';
import ColorScaleXform from './color-scale-xform.js';
import IconSetXform from './icon-set-xform.js';

const opType = (attributes: Record<string, string>): {type: string; operator?: string} => {
  const {type, operator} = attributes;
  switch (type) {
    case 'containsText':
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return {
        type: 'containsText',
        operator: type,
      };

    default:
      return {type, operator};
  }
};

class CfRuleXform extends CompositeXform<CfRuleModel> {
  databarXform: DatabarXform;
  extLstRefXform: ExtLstRefXform;
  formulaXform: FormulaXform;
  colorScaleXform: ColorScaleXform;
  iconSetXform: IconSetXform;

  constructor() {
    super();

    this.map = {
      dataBar: (this.databarXform = new DatabarXform()),
      extLst: (this.extLstRefXform = new ExtLstRefXform()),
      formula: (this.formulaXform = new FormulaXform()),
      colorScale: (this.colorScaleXform = new ColorScaleXform()),
      iconSet: (this.iconSetXform = new IconSetXform()),
    };
  }

  override tag = 'cfRule';

  override createNewModel(node?: XmlNode): CfRuleModel {
    const attributes = node?.attributes || {};
    return {
      ...opType(attributes),
      dxfId: BaseXform.toIntValue(attributes.dxfId),
      priority: BaseXform.toIntValue(attributes.priority),
      timePeriod: attributes.timePeriod,
      percent: BaseXform.toBoolValue(attributes.percent),
      bottom: BaseXform.toBoolValue(attributes.bottom),
      rank: BaseXform.toIntValue(attributes.rank),
      aboveAverage: BaseXform.toBoolValue(attributes.aboveAverage),
    };
  }

  override onParserClose(name: string, parser: {model: unknown}): void {
    switch (name) {
      case 'dataBar':
      case 'extLst':
      case 'colorScale':
      case 'iconSet':
        // merge parser model with ours
        Object.assign(this.model as CfRuleModel, parser.model);
        break;

      case 'formula':
        // except - formula is a string and appends to formulae
        {
          const m = this.model as CfRuleModel;
          m.formulae = m.formulae || [];
          m.formulae.push(parser.model as string);
        }
        break;
    }
  }
}

export default CfRuleXform;
export {CfRuleXform};
