import type {WorkbookViewModel} from '../../xform/book/workbook-view-xform.js';
export type {WorkbookViewModel} from '../../xform/book/workbook-view-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class WorkbookViewXform extends BaseXform<WorkbookViewModel> {
  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'workbookView') {
      const model = (this.model = {} as WorkbookViewModel);
      const addS = function (
        name: keyof WorkbookViewModel,
        value: string | undefined,
        dflt: string | undefined,
      ): void {
        const s = value !== undefined ? value : dflt;
        if (s !== undefined) {
          (model as Record<string, unknown>)[name] = s;
        }
      };
      const addN = function (
        name: keyof WorkbookViewModel,
        value: string | undefined,
        dflt: number | undefined,
      ): void {
        const n = value !== undefined ? parseInt(value, 10) : dflt;
        if (n !== undefined) {
          (model as Record<string, unknown>)[name] = n;
        }
      };
      addN('x', node.attributes.xWindow, 0);
      addN('y', node.attributes.yWindow, 0);
      addN('width', node.attributes.windowWidth, 25000);
      addN('height', node.attributes.windowHeight, 10000);
      addS('visibility', node.attributes.visibility, 'visible');
      addN('activeTab', node.attributes.activeTab, undefined);
      addN('firstSheet', node.attributes.firstSheet, undefined);
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default WorkbookViewXform;
export {WorkbookViewXform};
