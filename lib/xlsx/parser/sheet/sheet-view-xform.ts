import type {
  SelectionModel,
  PaneModel,
  SheetViewModel,
} from '../../xform/sheet/sheet-view-xform.js';
export type {
  SelectionModel,
  PaneModel,
  SheetViewModel,
} from '../../xform/sheet/sheet-view-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

const VIEW_STATES: Record<string, string> = {
  frozen: 'frozen',
  frozenSplit: 'frozen',
  split: 'split',
};

class SheetViewXform extends BaseXform<SheetViewModel> {
  sheetView?: SheetViewModel;
  pane?: PaneModel;
  selections!: Record<string, SelectionModel>;

  override tag = 'sheetView';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'sheetView':
        this.sheetView = {
          workbookViewId: parseInt(node.attributes.workbookViewId, 10),
          rightToLeft: node.attributes.rightToLeft === '1',
          tabSelected: node.attributes.tabSelected === '1',
          showRuler: !(node.attributes.showRuler === '0'),
          showRowColHeaders: !(node.attributes.showRowColHeaders === '0'),
          showGridLines: !(node.attributes.showGridLines === '0'),
          zoomScale: parseInt(node.attributes.zoomScale || '100', 10),
          zoomScaleNormal: parseInt(node.attributes.zoomScaleNormal || '100', 10),
          style: node.attributes.view,
        };
        this.pane = undefined;
        this.selections = {};
        return true;

      case 'pane':
        this.pane = {
          xSplit: parseInt(node.attributes.xSplit || '0', 10),
          ySplit: parseInt(node.attributes.ySplit || '0', 10),
          topLeftCell: node.attributes.topLeftCell,
          activePane: node.attributes.activePane || 'topLeft',
          state: node.attributes.state,
        };
        return true;

      case 'selection': {
        const name = node.attributes.pane || 'topLeft';
        this.selections[name] = {
          pane: name,
          activeCell: node.attributes.activeCell,
        };
        return true;
      }

      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    let model: SheetViewModel;
    let selection: SelectionModel | undefined;
    switch (name) {
      case 'sheetView':
        if (this.sheetView && this.pane) {
          model = this.model = {
            workbookViewId: this.sheetView.workbookViewId,
            rightToLeft: this.sheetView.rightToLeft,
            state: VIEW_STATES[this.pane.state as string] || 'split', // split is default
            xSplit: this.pane.xSplit,
            ySplit: this.pane.ySplit,
            topLeftCell: this.pane.topLeftCell,
            showRuler: this.sheetView.showRuler,
            showRowColHeaders: this.sheetView.showRowColHeaders,
            showGridLines: this.sheetView.showGridLines,
            zoomScale: this.sheetView.zoomScale,
            zoomScaleNormal: this.sheetView.zoomScaleNormal,
          };
          if (this.model.state === 'split') {
            model.activePane = this.pane.activePane;
          }
          selection = this.selections[this.pane.activePane];
          if (selection && selection.activeCell) {
            model.activeCell = selection.activeCell;
          }
          if (this.sheetView.style) {
            model.style = this.sheetView.style;
          }
        } else {
          model = this.model = {
            workbookViewId: this.sheetView!.workbookViewId,
            rightToLeft: this.sheetView!.rightToLeft,
            state: 'normal',
            showRuler: this.sheetView!.showRuler,
            showRowColHeaders: this.sheetView!.showRowColHeaders,
            showGridLines: this.sheetView!.showGridLines,
            zoomScale: this.sheetView!.zoomScale,
            zoomScaleNormal: this.sheetView!.zoomScaleNormal,
          };
          selection = this.selections.topLeft;
          if (selection && selection.activeCell) {
            model.activeCell = selection.activeCell;
          }
          if (this.sheetView!.style) {
            model.style = this.sheetView!.style;
          }
        }
        return false;
      default:
        return true;
    }
  }

  override reconcile(): void {}
}

export default SheetViewXform;
export {SheetViewXform};
