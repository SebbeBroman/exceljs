/** Geometry for compact cell ranges; work depends on rectangle count, never cell count. */
export interface Rectangle {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export function subtractRectangle(a: Rectangle, b: Rectangle): Rectangle[] {
  const top = Math.max(a.top, b.top),
    bottom = Math.min(a.bottom, b.bottom);
  const left = Math.max(a.left, b.left),
    right = Math.min(a.right, b.right);
  if (top > bottom || left > right) return [a];
  const pieces: Rectangle[] = [];
  if (a.top < top) pieces.push({...a, bottom: top - 1});
  if (bottom < a.bottom) pieces.push({...a, top: bottom + 1});
  if (a.left < left) pieces.push({top, bottom, left: a.left, right: left - 1});
  if (right < a.right) pieces.push({top, bottom, left: right + 1, right: a.right});
  return pieces;
}

function mergeAlong(input: Rectangle[], vertical: boolean): Rectangle[] {
  const spanStart = vertical ? 'left' : 'top';
  const spanEnd = vertical ? 'right' : 'bottom';
  const start = vertical ? 'top' : 'left';
  const end = vertical ? 'bottom' : 'right';
  input.sort(
    (a, b) => a[spanStart] - b[spanStart] || a[spanEnd] - b[spanEnd] || a[start] - b[start],
  );
  const result: Rectangle[] = [];
  for (const rect of input) {
    const previous = result[result.length - 1];
    if (
      previous &&
      previous[spanStart] === rect[spanStart] &&
      previous[spanEnd] === rect[spanEnd] &&
      rect[start] <= previous[end] + 1
    ) {
      previous[end] = Math.max(previous[end], rect[end]);
    } else result.push({...rect});
  }
  return result;
}

export function mergeRectangles(input: Rectangle[]): Rectangle[] {
  let result = input;
  let size: number;
  do {
    size = result.length;
    result = mergeAlong(mergeAlong(result, true), false);
  } while (result.length < size);
  return result.sort((a, b) => a.top - b.top || a.left - b.left);
}

export function unionRectangles(input: Rectangle[]): Rectangle[] {
  const result: Rectangle[] = [];
  // Sorted coalescing also deduplicates single cells; no pairwise subtraction is needed.
  if (input.every(rect => rect.top === rect.bottom && rect.left === rect.right)) {
    return mergeRectangles(input);
  }
  for (const rect of mergeRectangles(input)) {
    let pieces = [rect];
    for (const existing of result)
      pieces = pieces.flatMap(piece => subtractRectangle(piece, existing));
    result.push(...pieces);
  }
  return mergeRectangles(result);
}
