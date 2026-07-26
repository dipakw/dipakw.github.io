/** 0-based column index -> spreadsheet column name (0 -> "A", 26 -> "AA"). */
export function colName(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = (c - m - 1) / 26;
  }
  return s;
}

/** Column name -> 0-based column index ("A" -> 0, "AA" -> 26). */
export function colIndex(s: string): number {
  let n = 0;
  const u = s.toUpperCase();
  for (let i = 0; i < u.length; i++) n = n * 26 + (u.charCodeAt(i) - 64);
  return n - 1;
}

/**
 * Coerce a raw cell string to a number, or null if it isn't numeric.
 * Accepts thousands separators and a trailing % (which divides by 100).
 */
export function numeric(v: unknown): number | null {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return null;
  const t = v.trim().replace(/,/g, '');
  if (t === '' || !/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?%?$/i.test(t)) return null;
  return t.endsWith('%') ? parseFloat(t) / 100 : parseFloat(t);
}
