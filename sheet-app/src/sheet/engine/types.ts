/** Horizontal text alignment for a cell. */
export type Align = 'left' | 'center' | 'right';

/**
 * Number display format.
 * '' / undefined = automatic (the engine picks a sensible representation).
 */
export type NumberFormat = 'int' | 'n2' | 'k' | 'cur' | 'pct' | '';

/**
 * Custom per-side cell borders. Each side holds a CSS color when present;
 * absent = no custom border on that side (the gridline, if any, shows instead).
 */
export interface CellBorders {
  t?: string;
  r?: string;
  b?: string;
  l?: string;
}

/** Which borders an action targets within the selection. */
export type BorderMode = 'all' | 'outer' | 'top' | 'right' | 'bottom' | 'left' | 'none';

/**
 * The stored model for a single cell. Everything is optional so an empty cell
 * costs nothing — the engine prunes cells that carry no value and no formatting.
 *
 * `v` is the RAW user input: a literal ("42", "Revenue") or a formula ("=SUM(A1:A9)").
 * The computed/displayed value is derived on demand by the engine, never stored.
 */
export interface CellData {
  /** Raw value or formula source (formulas start with '='). */
  v?: string;
  /** Bold. */
  b?: 0 | 1;
  /** Italic. */
  i?: 0 | 1;
  /** Strikethrough. */
  s?: 0 | 1;
  /** Font family (CSS font-family stack); '' / undefined = default. */
  ff?: string;
  /** Text color (CSS color); '' / undefined = default ink. */
  co?: string;
  /** Alignment override. */
  a?: Align;
  /** Number format. */
  f?: NumberFormat;
  /** Background fill (CSS color); '' / undefined = no fill. */
  g?: string;
  /** Custom per-side borders. */
  bd?: CellBorders;
}

/** Sparse cell store keyed by `"row,col"`. */
export type CellMap = Record<string, CellData>;

/** A fully-evaluated cell value. */
export type CellValue = number | string | boolean;

/** Serializable sheet state — this is exactly what a StorageAdapter persists. */
export interface SheetSnapshot {
  data: CellMap;
  title: string;
  /** Grid dimensions (grow over time). Absent = engine defaults. */
  rows?: number;
  cols?: number;
  /** Per-column pixel widths, keyed by column index. Absent = default width. */
  colWidths?: Record<number, number>;
  /** Whether the default gridlines are shown. Absent = true. */
  gridlines?: boolean;
  /** Merged cell regions. */
  merges?: MergeRegion[];
  /** Number of frozen rows (kept visible at the top) / columns (at the left). */
  frozenRows?: number;
  frozenCols?: number;
}

/** Live selection: anchor (r,c) + the opposite corner (r2,c2). */
export interface Selection {
  r: number;
  c: number;
  r2: number;
  c2: number;
}

/** A normalized rectangle (min/max already resolved). */
export interface Range {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/** A merged cell region. The top-left (r1,c1) holds the content. */
export interface MergeRegion {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Internal error signal thrown/returned by the formula evaluator. */
export interface FormulaError {
  e: string;
}
