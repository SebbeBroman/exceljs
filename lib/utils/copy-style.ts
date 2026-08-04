/** Loose style-like object (partial nested style tree). */
type StyleLike = Record<string, unknown>;

const oneDepthCopy = (obj: StyleLike, nestKeys: string[]): StyleLike => ({
  ...obj,
  ...nestKeys.reduce<StyleLike>((memo, key) => {
    if (obj[key]) memo[key] = {...(obj[key] as StyleLike)};
    return memo;
  }, {}),
});

const setIfExists = (
  src: StyleLike,
  dst: StyleLike,
  key: string,
  nestKeys: string[] = [],
): void => {
  if (src[key]) dst[key] = oneDepthCopy(src[key] as StyleLike, nestKeys);
};

const isEmptyObj = (obj: object): boolean => Object.keys(obj).length === 0;

const copyStyle = <T extends StyleLike | null | undefined>(style: T): T | StyleLike => {
  if (!style) return style;
  if (isEmptyObj(style)) return {};

  const copied: StyleLike = {...style};

  setIfExists(style, copied, 'font', ['color']);
  setIfExists(style, copied, 'alignment');
  setIfExists(style, copied, 'protection');
  if (style.border) {
    setIfExists(style, copied, 'border');
    const border = style.border as StyleLike;
    const copiedBorder = copied.border as StyleLike;
    setIfExists(border, copiedBorder, 'top', ['color']);
    setIfExists(border, copiedBorder, 'left', ['color']);
    setIfExists(border, copiedBorder, 'bottom', ['color']);
    setIfExists(border, copiedBorder, 'right', ['color']);
    setIfExists(border, copiedBorder, 'diagonal', ['color']);
  }

  if (style.fill) {
    setIfExists(style, copied, 'fill', ['fgColor', 'bgColor', 'center']);
    const fill = style.fill as StyleLike & {stops?: StyleLike[]};
    if (fill.stops) {
      (copied.fill as StyleLike).stops = fill.stops.map(s => oneDepthCopy(s, ['color']));
    }
  }

  return copied;
};
export {copyStyle};
