export function xmlDecode(text: string): string {
  return text.replace(/&([a-z]*);/g, c => {
    switch (c) {
      case '&lt;':
        return '<';
      case '&gt;':
        return '>';
      case '&amp;':
        return '&';
      case '&apos;':
        return "'";
      case '&quot;':
        return '"';
      default:
        return c;
    }
  });
}
