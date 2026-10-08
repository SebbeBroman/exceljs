export {xmlDecode} from './xml-decode.js';

// oxlint-disable-next-line no-control-regex
const xmlDecodeRegex = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;
export function dateToExcel(d: Date, date1904?: boolean): number {
  return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
}

export function excelToDate(v: number, date1904?: boolean): Date {
  const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
  return new Date(millisecondSinceEpoch);
}

export function xmlEncode(text: string): string {
  const regexResult = xmlDecodeRegex.exec(text);
  if (!regexResult) return text;

  let result = '';
  let escape = '';
  let lastIndex = 0;
  let i = regexResult.index;
  for (; i < text.length; i++) {
    const charCode = text.charCodeAt(i);
    switch (charCode) {
      case 34: // "
        escape = '&quot;';
        break;
      case 38: // &
        escape = '&amp;';
        break;
      case 39: // '
        escape = '&apos;';
        break;
      case 60: // <
        escape = '&lt;';
        break;
      case 62: // >
        escape = '&gt;';
        break;
      case 127:
        escape = '';
        break;
      default: {
        if (charCode <= 31 && (charCode <= 8 || (charCode >= 11 && charCode !== 13))) {
          escape = '';
          break;
        }
        continue;
      }
    }
    if (lastIndex !== i) result += text.substring(lastIndex, i);
    lastIndex = i + 1;
    if (escape) result += escape;
  }
  if (lastIndex !== i) return result + text.substring(lastIndex, i);
  return result;
}

export function validInt(value: unknown): number {
  const i = parseInt(String(value), 10);
  return !Number.isNaN(i) ? i : 0;
}

export function isDateFmt(fmt: string | null | undefined): boolean {
  if (!fmt) {
    return false;
  }

  // must remove all chars inside quotes and []
  fmt = fmt.replace(/\[[^\]]*]/g, '');
  fmt = fmt.replace(/"[^"]*"/g, '');
  // then check for date formatting chars
  const result = fmt.match(/[ymdhMsb]+/) !== null;
  return result;
}

export function parseBoolean(value: unknown): boolean {
  return value === true || value === 'true' || value === 1 || value === '1';
}
