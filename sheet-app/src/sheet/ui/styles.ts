import type { CSSProperties } from 'react';

/** `color-mix` helper: blend `pct`% of the accent into a base color. */
export function mix(accent: string, pct: number, base: string): string {
  return `color-mix(in oklab, ${accent} ${pct}%, ${base})`;
}

/** Neutral bordered action button (insert rows/columns). */
export const neutralButton: CSSProperties = {
  padding: '6px 10px',
  borderRadius: 5,
  border: '1px solid var(--ex-border-2)',
  background: 'var(--ex-surface)',
  color: 'var(--ex-text-muted)',
  font: "500 12.5px/1 var(--ex-font-sans)",
  cursor: 'pointer',
  outline: 'none',
};

/** Red text button for a destructive action (clear selection). */
export const dangerGhost: CSSProperties = {
  padding: '6px 10px',
  borderRadius: 5,
  border: '1px solid transparent',
  background: 'transparent',
  color: 'var(--ex-danger)',
  font: "500 12.5px/1 var(--ex-font-sans)",
  cursor: 'pointer',
  outline: 'none',
};

/** Red bordered button for a destructive action (reset all). */
export const dangerButton: CSSProperties = {
  padding: '6px 10px',
  borderRadius: 5,
  border: '1px solid var(--ex-danger-border)',
  background: 'var(--ex-surface)',
  color: 'var(--ex-danger)',
  font: "500 12.5px/1 var(--ex-font-sans)",
  cursor: 'pointer',
  outline: 'none',
};
