import type {CellValue} from '../model/types.js';

/** Stringify a cell for grid UIs (SheetJS `raw: false`-ish). */
export function cellToDisplayString(value: CellValue, defval = ''): string {
  if (value == null || value === '') return defval;

  if (typeof value === 'string') return value;
  if (typeof value === 'number') {
    if (Number.isNaN(value)) return defval;
    return String(value);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return defval;
    return value.toISOString();
  }

  if (typeof value === 'object') {
    if ('formula' in value || 'sharedFormula' in value) {
      const result = (value as {result?: unknown}).result;
      if (result !== undefined) return cellToDisplayString(result as CellValue, defval);
      return defval;
    }
    if ('text' in value || 'hyperlink' in value) {
      const t = (value as {text?: unknown}).text;
      if (t != null) return cellToDisplayString(t as CellValue, defval);
      return String((value as {hyperlink?: string}).hyperlink ?? defval);
    }
    if ('richText' in value && Array.isArray((value as {richText: {text?: string}[]}).richText)) {
      return (value as {richText: {text?: string}[]}).richText.map(r => r.text ?? '').join('');
    }
    if ('error' in value) {
      return String((value as {error: string}).error ?? defval);
    }
  }

  return String(value);
}
