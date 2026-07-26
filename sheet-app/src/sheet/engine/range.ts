import type { Range, Selection } from './types';

/** Resolve a selection (anchor + opposite corner) into a normalized rectangle. */
export function selectionToRange(s: Selection): Range {
  return {
    r1: Math.min(s.r, s.r2),
    r2: Math.max(s.r, s.r2),
    c1: Math.min(s.c, s.c2),
    c2: Math.max(s.c, s.c2),
  };
}

/** Iterate every (row, col) inside a rectangle. */
export function forEachInRange(range: Range, fn: (r: number, c: number) => void): void {
  for (let r = range.r1; r <= range.r2; r++) {
    for (let c = range.c1; c <= range.c2; c++) fn(r, c);
  }
}
