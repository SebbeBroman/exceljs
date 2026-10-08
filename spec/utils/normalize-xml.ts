export function normalizeXml(xml: string): string {
  let s = String(xml)
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .replace(/>\s+/g, '>')
    .replace(/\s+</g, '<')
    .replace(/\s+\/>/g, '/>') // <tag /> → <tag/>
    .trim();
  // Sort attributes so order differences don't fail (chai-xml was order-tolerant)
  s = s.replace(
    /<([A-Za-z0-9:_-]+)((?:\s+[A-Za-z0-9:_-]+="[^"]*")*)(\s*\/?)>/g,
    (_m, tag, attrs, close) => {
      const selfClose = close.includes('/') ? '/' : '';
      if (!attrs || !attrs.trim()) return `<${tag}${selfClose}>`;
      const list: string[] = [];
      const re = /([A-Za-z0-9:_-]+)="([^"]*)"/g;
      let m;
      while ((m = re.exec(attrs))) {
        list.push(`${m[1]}="${m[2]}"`);
      }
      list.sort();
      return `<${tag}${list.length ? ` ${list.join(' ')}` : ''}${selfClose}>`;
    },
  );
  // <tag attrs></tag> with no content → <tag attrs/>
  s = s.replace(/<([A-Za-z0-9:_-]+)((?:\s+[A-Za-z0-9:_-]+="[^"]*")*)><\/\1>/g, '<$1$2/>');
  return s;
}
