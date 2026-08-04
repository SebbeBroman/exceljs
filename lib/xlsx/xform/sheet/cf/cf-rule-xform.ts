import BaseXform from '../../base-xform.js';
import CompositeXform from '../../composite-xform.js';
import type {XmlStreamLike, XmlNode} from '../../base-xform.js';
import Range from '../../../../doc/range.js';
import DatabarXform from './databar-xform.js';
import ExtLstRefXform from './ext-lst-ref-xform.js';
import FormulaXform from './formula-xform.js';
import ColorScaleXform from './color-scale-xform.js';
import IconSetXform from './icon-set-xform.js';

const extIcons: Record<string, boolean> = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

export interface CfRuleModel {
  type?: string;
  operator?: string;
  formulae?: string[];
  ref?: string;
  text?: string;
  timePeriod?: string;
  dxfId?: number;
  priority?: number;
  percent?: boolean;
  bottom?: boolean;
  rank?: number;
  aboveAverage?: boolean;
  style?: unknown;
  iconSet?: string;
  custom?: boolean;
  x14Id?: string;
  cfvo?: unknown[];
  color?: unknown;
  [key: string]: unknown;
}

const getTextFormula = (model: CfRuleModel): string | undefined => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.operator) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH("${model.text}",${tl})))`;
    case 'containsBlanks':
      return `LEN(TRIM(${tl}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${tl}))>0`;
    case 'containsErrors':
      return `ISERROR(${tl})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${tl}))`;
    default:
      return undefined;
  }
};

const getTimePeriodFormula = (model: CfRuleModel): string | undefined => {
  if (model.formulae && model.formulae[0]) {
    return model.formulae[0];
  }

  const range = new Range(model.ref);
  const {tl} = range;
  switch (model.timePeriod) {
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case 'nextWeek':
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case 'yesterday':
      return `FLOOR(${tl},1)=TODAY()-1`;
    case 'today':
      return `FLOOR(${tl},1)=TODAY()`;
    case 'tomorrow':
      return `FLOOR(${tl},1)=TODAY()+1`;
    case 'last7Days':
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case 'lastMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case 'thisMonth':
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case 'nextMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
};

const opType = (
  attributes: Record<string, string>,
): {type: string; operator?: string} => {
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

  static isPrimitive(rule: CfRuleModel): boolean {
    // is this rule primitive?
    if (rule.type === 'iconSet') {
      if (rule.custom || (rule.iconSet && extIcons[rule.iconSet])) {
        return false;
      }
    }
    return true;
  }

  override render(xmlStream: XmlStreamLike, model?: CfRuleModel | null): void {
    if (!model) {
      return;
    }
    switch (model.type) {
      case 'expression':
        this.renderExpression(xmlStream, model);
        break;
      case 'cellIs':
        this.renderCellIs(xmlStream, model);
        break;
      case 'top10':
        this.renderTop10(xmlStream, model);
        break;
      case 'aboveAverage':
        this.renderAboveAverage(xmlStream, model);
        break;
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'colorScale':
        this.renderColorScale(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
      case 'containsText':
        this.renderText(xmlStream, model);
        break;
      case 'timePeriod':
        this.renderTimePeriod(xmlStream, model);
        break;
    }
  }

  renderExpression(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'expression',
      dxfId: model.dxfId,
      priority: model.priority,
    });

    this.formulaXform.render(xmlStream, model.formulae![0]);

    xmlStream.closeNode();
  }

  renderCellIs(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'cellIs',
      dxfId: model.dxfId,
      priority: model.priority,
      operator: model.operator,
    });

    model.formulae!.forEach(formula => {
      this.formulaXform.render(xmlStream, formula);
    });

    xmlStream.closeNode();
  }

  renderTop10(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.leafNode(this.tag, {
      type: 'top10',
      dxfId: model.dxfId,
      priority: model.priority,
      percent: BaseXform.toBoolAttribute(model.percent, false),
      bottom: BaseXform.toBoolAttribute(model.bottom, false),
      // original called toIntValue(model.rank, 10, true) — third arg unused by toIntValue
      rank: BaseXform.toIntValue(model.rank as unknown as string | undefined, 10),
    });
  }

  renderAboveAverage(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.leafNode(this.tag, {
      type: 'aboveAverage',
      dxfId: model.dxfId,
      priority: model.priority,
      aboveAverage: BaseXform.toBoolAttribute(model.aboveAverage, true),
    });
  }

  renderDataBar(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      priority: model.priority,
    });

    this.databarXform.render(xmlStream, model as never);
    this.extLstRefXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderColorScale(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'colorScale',
      priority: model.priority,
    });

    this.colorScaleXform.render(xmlStream, model as never);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    // iconset is all primitive or all extLst
    if (!CfRuleXform.isPrimitive(model)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
    });

    this.iconSetXform.render(xmlStream, model as never);

    xmlStream.closeNode();
  }

  renderText(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: model.operator,
      dxfId: model.dxfId,
      priority: model.priority,
      operator: BaseXform.toStringAttribute(model.operator, 'containsText'),
    });

    const formula = getTextFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  renderTimePeriod(xmlStream: XmlStreamLike, model: CfRuleModel): void {
    xmlStream.openNode(this.tag, {
      type: 'timePeriod',
      dxfId: model.dxfId,
      priority: model.priority,
      timePeriod: model.timePeriod,
    });

    const formula = getTimePeriodFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

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
