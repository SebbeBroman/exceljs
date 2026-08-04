const escapeHtmlRegex = /["&<>]/;

/** Escape `"`, `&`, `'`, `<`, `>` for HTML text content. */
export function escapeHtml(html: string): string {
  const regexResult = escapeHtmlRegex.exec(html);
  if (!regexResult) return html;

  let result = '';
  let escape = '';
  let lastIndex = 0;
  let i = regexResult.index;
  for (; i < html.length; i++) {
    switch (html.charAt(i)) {
      case '"':
        escape = '&quot;';
        break;
      case '&':
        escape = '&amp;';
        break;
      case "'":
        escape = '&apos;';
        break;
      case '<':
        escape = '&lt;';
        break;
      case '>':
        escape = '&gt;';
        break;
      default:
        continue;
    }
    if (lastIndex !== i) result += html.substring(lastIndex, i);
    lastIndex = i + 1;
    result += escape;
  }
  if (lastIndex !== i) return result + html.substring(lastIndex, i);
  return result;
}
