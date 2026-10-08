/** Shared normalization for public workbook snapshots. */
import type {Style, WorkbookMeta} from './types.js';

export function pickStyle(style: Record<string, unknown> | undefined | null): Style | undefined {
  if (!style || typeof style !== 'object') return undefined;
  const out: Style = {};
  if (style.numFmt != null) out.numFmt = style.numFmt as string;
  if (style.font) out.font = style.font as Style['font'];
  if (style.alignment) out.alignment = style.alignment as Style['alignment'];
  if (style.protection) out.protection = style.protection as Style['protection'];
  if (style.border) out.border = style.border as Style['border'];
  if (style.fill) out.fill = style.fill as Style['fill'];
  return Object.keys(out).length ? out : undefined;
}

export function metaFromModel(wb: WorkbookMeta): WorkbookMeta {
  const meta: WorkbookMeta = {};
  if (wb.creator != null && wb.creator !== '') meta.creator = wb.creator;
  if (wb.lastModifiedBy != null && wb.lastModifiedBy !== '')
    meta.lastModifiedBy = wb.lastModifiedBy;
  if (wb.created) meta.created = wb.created;
  if (wb.modified) meta.modified = wb.modified;
  if (wb.company) meta.company = wb.company;
  if (wb.manager) meta.manager = wb.manager;
  if (wb.title) meta.title = wb.title;
  if (wb.subject) meta.subject = wb.subject;
  if (wb.keywords) meta.keywords = wb.keywords;
  if (wb.category) meta.category = wb.category;
  if (wb.description) meta.description = wb.description;
  if (wb.language) meta.language = wb.language;
  if (wb.revision != null) meta.revision = wb.revision;
  if (wb.contentStatus) meta.contentStatus = wb.contentStatus;
  if (wb.properties && Object.keys(wb.properties).length) {
    meta.properties = wb.properties as WorkbookMeta['properties'];
  }
  if (wb.views && wb.views.length) meta.views = wb.views;
  return meta;
}
