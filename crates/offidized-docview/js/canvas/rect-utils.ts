export type RectLTRB = [number, number, number, number];

export function extractRectLTRB(value: unknown): RectLTRB | null {
  if (!value) return null;

  if (Array.isArray(value)) {
    if (value.length < 4) return null;
    const left = Number(value[0]);
    const top = Number(value[1]);
    const right = Number(value[2]);
    const bottom = Number(value[3]);
    return Number.isFinite(left) &&
      Number.isFinite(top) &&
      Number.isFinite(right) &&
      Number.isFinite(bottom)
      ? [left, top, right, bottom]
      : null;
  }

  if (
    ArrayBuffer.isView(value) &&
    typeof value === "object" &&
    "length" in value
  ) {
    const arr = value as unknown as { length: number; [index: number]: number };
    if (arr.length < 4) return null;
    const left = Number(arr[0]);
    const top = Number(arr[1]);
    const right = Number(arr[2]);
    const bottom = Number(arr[3]);
    return Number.isFinite(left) &&
      Number.isFinite(top) &&
      Number.isFinite(right) &&
      Number.isFinite(bottom)
      ? [left, top, right, bottom]
      : null;
  }

  if (typeof value === "object") {
    const rec = value as Record<string, unknown>;
    const rectField = rec.rect;
    if (rectField) {
      return extractRectLTRB(rectField);
    }
    const left = Number(rec.left ?? rec.x);
    const top = Number(rec.top ?? rec.y);
    const right = Number(
      rec.right ??
        (typeof rec.w === "number" && Number.isFinite(left)
          ? left + Number(rec.w)
          : NaN),
    );
    const bottom = Number(
      rec.bottom ??
        (typeof rec.h === "number" && Number.isFinite(top)
          ? top + Number(rec.h)
          : NaN),
    );
    return Number.isFinite(left) &&
      Number.isFinite(top) &&
      Number.isFinite(right) &&
      Number.isFinite(bottom)
      ? [left, top, right, bottom]
      : null;
  }

  return null;
}

export function extractRectList(value: unknown): RectLTRB[] {
  if (!value) return [];

  if (Array.isArray(value)) {
    if (value.length === 0) return [];
    if (typeof value[0] === "number") {
      const single = extractRectLTRB(value);
      return single ? [single] : [];
    }
    const out: RectLTRB[] = [];
    for (const item of value) {
      const rect = extractRectLTRB(item);
      if (rect) out.push(rect);
    }
    return out;
  }

  const single = extractRectLTRB(value);
  return single ? [single] : [];
}

export function getFirstRangeRect(
  paragraph: any,
  ck: any,
  start: number,
  end: number,
): RectLTRB | null {
  const raw = paragraph.getRectsForRange(
    start,
    end,
    ck.RectHeightStyle.Max,
    ck.RectWidthStyle.Tight,
  ) as unknown;
  const rects = extractRectList(raw);
  return rects[0] ?? null;
}
