import colCache from '../../../utils/col-cache.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface SheetViewModel {
  workbookViewId?: number;
  rightToLeft?: boolean;
  tabSelected?: boolean;
  showRuler?: boolean;
  showRowColHeaders?: boolean;
  showGridLines?: boolean;
  zoomScale?: number;
  zoomScaleNormal?: number;
  style?: string;
  state?: string;
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activePane?: string;
  activeCell?: string;
}

export interface PaneModel {
  xSplit: number;
  ySplit: number;
  topLeftCell?: string;
  activePane: string;
  state?: string;
}

export interface SelectionModel {
  pane: string;
  activeCell?: string;
}

class SheetViewXform extends BaseXform<SheetViewModel> {
  sheetView?: SheetViewModel;
  pane?: PaneModel;
  selections!: Record<string, SelectionModel>;

  override tag = 'sheetView';

  override prepare(model?: SheetViewModel | null): void {
    if (!model) {
      return;
    }
    switch (model.state) {
      case 'frozen':
      case 'split':
        break;
      default:
        model.state = 'normal';
        break;
    }
  }

  override render(xmlStream: XmlStreamLike, model?: SheetViewModel | null): void {
    if (!model) {
      return;
    }
    xmlStream.openNode('sheetView', {
      workbookViewId: model.workbookViewId || 0,
    });
    const add = function (name: string, value: unknown, included: unknown): void {
      if (included) {
        xmlStream.addAttribute(name, value);
      }
    };
    add('rightToLeft', '1', model.rightToLeft === true);
    add('tabSelected', '1', model.tabSelected);
    add('showRuler', '0', model.showRuler === false);
    add('showRowColHeaders', '0', model.showRowColHeaders === false);
    add('showGridLines', '0', model.showGridLines === false);
    add('zoomScale', model.zoomScale, model.zoomScale);
    add('zoomScaleNormal', model.zoomScaleNormal, model.zoomScaleNormal);
    add('view', model.style, model.style);

    let topLeftCell: string;
    let xSplit: number;
    let ySplit: number;
    let activePane: string;
    switch (model.state) {
      case 'frozen':
        xSplit = model.xSplit || 0;
        ySplit = model.ySplit || 0;
        topLeftCell = model.topLeftCell || colCache.getAddress(ySplit + 1, xSplit + 1).address;
        activePane =
          (model.xSplit && model.ySplit && 'bottomRight') ||
          (model.xSplit && 'topRight') ||
          'bottomLeft';

        xmlStream.leafNode('pane', {
          xSplit: model.xSplit || undefined,
          ySplit: model.ySplit || undefined,
          topLeftCell,
          activePane,
          state: 'frozen',
        });
        xmlStream.leafNode('selection', {
          pane: activePane,
          activeCell: model.activeCell,
          sqref: model.activeCell,
        });
        break;
      case 'split':
        if (model.activePane === 'topLeft') {
          model.activePane = undefined;
        }
        xmlStream.leafNode('pane', {
          xSplit: model.xSplit || undefined,
          ySplit: model.ySplit || undefined,
          topLeftCell: model.topLeftCell,
          activePane: model.activePane,
        });
        xmlStream.leafNode('selection', {
          pane: model.activePane,
          activeCell: model.activeCell,
          sqref: model.activeCell,
        });
        break;
      case 'normal':
        if (model.activeCell) {
          xmlStream.leafNode('selection', {
            activeCell: model.activeCell,
            sqref: model.activeCell,
          });
        }
        break;
      default:
        break;
    }
    xmlStream.closeNode();
  }
}

export default SheetViewXform;
export {SheetViewXform};
