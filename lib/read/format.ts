/** Format sniffing shared by viewWorkbook / readRows. */

export type ViewFormat = 'auto' | 'csv' | 'xlsx';

export function toUint8Array(data: ArrayBuffer | Uint8Array | ArrayBufferView): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}

export function decodeText(data: ArrayBuffer | Uint8Array | string, encoding: string): string {
  if (typeof data === 'string') return data;
  return new TextDecoder(encoding).decode(toUint8Array(data));
}

/**
 * Resolve csv vs OOXML (xlsx/xlsm/…).
 * `.xls` / `.xlsb` are not supported — callers get a clear error if forced wrong.
 */
export function sniffFormat(
  data: ArrayBuffer | Uint8Array | string,
  filename?: string,
  format?: ViewFormat,
): 'csv' | 'xlsx' {
  if (format === 'csv' || format === 'xlsx') return format;

  const lower = (filename || '').toLowerCase();
  if (lower.endsWith('.csv') || lower.endsWith('.tsv') || lower.endsWith('.txt')) {
    return 'csv';
  }
  if (
    lower.endsWith('.xlsx') ||
    lower.endsWith('.xlsm') ||
    lower.endsWith('.xltx') ||
    lower.endsWith('.xltm')
  ) {
    return 'xlsx';
  }
  if (lower.endsWith('.xls') || lower.endsWith('.xlsb')) {
    throw new Error(
      `Unsupported format "${lower.slice(lower.lastIndexOf('.'))}": only OOXML (.xlsx/.xlsm/…) and CSV are supported`,
    );
  }

  if (typeof data !== 'string') {
    const u8 = toUint8Array(data);
    if (u8.length >= 2 && u8[0] === 0x50 && u8[1] === 0x4b) {
      return 'xlsx';
    }
  }

  return 'csv';
}
