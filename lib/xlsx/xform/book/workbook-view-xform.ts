import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface WorkbookViewModel {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  firstSheet?: number;
  activeTab?: number;
  visibility?: string;
}

class WorkbookViewXform extends BaseXform<WorkbookViewModel> {
  override render(xmlStream: XmlStreamLike, model?: WorkbookViewModel | null): void {
    const attributes: Record<string, unknown> = {
      xWindow: model!.x || 0,
      yWindow: model!.y || 0,
      windowWidth: model!.width || 12000,
      windowHeight: model!.height || 24000,
      firstSheet: model!.firstSheet,
      activeTab: model!.activeTab,
    };
    if (model!.visibility && model!.visibility !== 'visible') {
      attributes.visibility = model!.visibility;
    }
    xmlStream.leafNode('workbookView', attributes);
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'workbookView') {
      const model = (this.model = {} as WorkbookViewModel);
      const addS = function (name: keyof WorkbookViewModel, value: string | undefined, dflt: string | undefined): void {
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
