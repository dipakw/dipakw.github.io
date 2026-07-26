import type { CellValue, NumberFormat } from './types';

/**
 * Render an evaluated value for display given a number format.
 * Strings and booleans pass through untouched; only numbers are formatted.
 */
export function format(v: CellValue | '' | null | undefined, f?: NumberFormat): string {
  if (v === '' || v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;

  const n = v;
  if (f === 'int') return Math.round(n).toLocaleString('en-US');
  if (f === 'n2') return n.toFixed(2);
  if (f === 'k') return n.toLocaleString('en-US', { maximumFractionDigits: 2 });
  if (f === 'cur') {
    return (n < 0 ? '-$' : '$') + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  if (f === 'pct') return (n * 100).toLocaleString('en-US', { maximumFractionDigits: 2 }) + '%';

  // Automatic: trim floating-point noise, fall back to precision for long values.
  const r = Math.round(n * 1e10) / 1e10;
  return String(r).length > 13 ? r.toPrecision(10).replace(/\.?0+$/, '') : String(r);
}
