import type { DocSelection } from "../adapter.ts";

export interface NormalizedSelectionRange {
  startBody: number;
  startChar: number;
  endBody: number;
  endChar: number;
}

export function normalizeSelectionRange(
  selection: DocSelection,
): NormalizedSelectionRange | null {
  let startBody = selection.anchor.bodyIndex;
  let startChar = selection.anchor.charOffset;
  let endBody = selection.focus.bodyIndex;
  let endChar = selection.focus.charOffset;

  if (startBody > endBody || (startBody === endBody && startChar > endChar)) {
    [startBody, startChar, endBody, endChar] = [
      endBody,
      endChar,
      startBody,
      startChar,
    ];
  }

  if (startBody === endBody && startChar === endChar) return null;

  return { startBody, startChar, endBody, endChar };
}

export function selectionRangeInParagraphFragment(
  normalized: NormalizedSelectionRange,
  bodyIndex: number,
  utf16Start: number,
  utf16End: number,
  paragraphEnd: number,
): { start: number; end: number } | null {
  if (bodyIndex < normalized.startBody || bodyIndex > normalized.endBody) {
    return null;
  }

  let start = 0;
  let end = paragraphEnd;
  if (bodyIndex === normalized.startBody) start = normalized.startChar;
  if (bodyIndex === normalized.endBody) end = normalized.endChar;

  const fragmentEnd = Math.max(utf16End, utf16Start);
  start = Math.max(start, utf16Start);
  end = Math.min(end, fragmentEnd);
  if (start >= end) return null;

  return { start, end };
}
