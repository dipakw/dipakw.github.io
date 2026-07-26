import { colIndex, colName, numeric } from './cellref';
import { evaluate, rewriteFormulaRefs } from './formula';
import { forEachInRange } from './range';
import type {
  BorderMode,
  CellBorders,
  CellData,
  CellMap,
  CellValue,
  FormulaError,
  MergeRegion,
  Range,
  SheetSnapshot,
} from './types';

/** Full mutable engine state — one history entry restores all of it. */
interface EngineState {
  data: CellMap;
  rows: number;
  cols: number;
  colWidths: Record<number, number>;
  gridlines: boolean;
  merges: MergeRegion[];
  frozenRows: number;
  frozenCols: number;
}

export interface EngineOptions {
  rows?: number;
  cols?: number;
  /** Initial state to hydrate from (e.g. loaded from storage). */
  snapshot?: SheetSnapshot | null;
  defaultTitle?: string;
  /** Fallback column width (px) for columns without an explicit size. */
  defaultColWidth?: number;
}

export const MIN_COL_WIDTH = 40;
export const MAX_COL_WIDTH = 640;
export const MAX_ROWS = 1000;
export const MAX_COLS = 256;

export interface RangeStats {
  sum: number;
  count: number;
  filled: number;
}

const DEFAULT_TITLE = 'Untitled draft';

/**
 * Framework-agnostic spreadsheet core: the data model, the formula engine,
 * undo/redo history, and a tiny pub/sub so any UI (React today, anything later)
 * can subscribe to changes.
 *
 * It knows nothing about the DOM, React, or where data is persisted — storage is
 * handled outside via a StorageAdapter + the useSpreadsheet hook.
 */
export class SpreadsheetEngine {
  readonly defaultColWidth: number;
  private readonly defaultRows: number;
  private readonly defaultCols: number;

  private _rows: number;
  private _cols: number;
  private data: CellMap;
  private title: string;
  private colWidths: Record<number, number>;
  private gridlines: boolean;
  private merges: MergeRegion[];
  private frozenRows: number;
  private frozenCols: number;
  private cache: Record<string, CellValue> = {};
  private hist: EngineState[] = [];
  private future: EngineState[] = [];
  private listeners = new Set<() => void>();

  constructor(opts: EngineOptions = {}) {
    this.defaultRows = Math.max(10, opts.rows ?? 60);
    this.defaultCols = Math.max(4, opts.cols ?? 26);
    this.defaultColWidth = opts.defaultColWidth ?? 104;
    this._rows = Math.min(MAX_ROWS, opts.snapshot?.rows ?? this.defaultRows);
    this._cols = Math.min(MAX_COLS, opts.snapshot?.cols ?? this.defaultCols);
    this.data = opts.snapshot?.data ? { ...opts.snapshot.data } : {};
    this.title = opts.snapshot?.title ?? opts.defaultTitle ?? DEFAULT_TITLE;
    this.colWidths = { ...(opts.snapshot?.colWidths ?? {}) };
    this.gridlines = opts.snapshot?.gridlines ?? true;
    this.merges = (opts.snapshot?.merges ?? []).map((m) => ({ ...m }));
    this.frozenRows = opts.snapshot?.frozenRows ?? 0;
    this.frozenCols = opts.snapshot?.frozenCols ?? 0;
  }

  /* ---------- full-state history helpers ---------- */
  private capture(): EngineState {
    return {
      data: { ...this.data },
      rows: this._rows,
      cols: this._cols,
      colWidths: { ...this.colWidths },
      gridlines: this.gridlines,
      merges: this.merges.map((m) => ({ ...m })),
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
    };
  }
  private restore(s: EngineState): void {
    this.data = { ...s.data };
    this._rows = s.rows;
    this._cols = s.cols;
    this.colWidths = { ...s.colWidths };
    this.gridlines = s.gridlines;
    this.merges = s.merges.map((m) => ({ ...m }));
    this.frozenRows = s.frozenRows;
    this.frozenCols = s.frozenCols;
    this.cache = {};
  }
  /** Snapshot current state onto the undo stack (before an out-of-band change). */
  checkpoint(): void {
    this.hist = this.hist.concat([this.capture()]).slice(-80);
    this.future = [];
  }

  get rows(): number {
    return this._rows;
  }
  get cols(): number {
    return this._cols;
  }

  /* ---------- subscription ---------- */

  /** Subscribe to any state change. Returns an unsubscribe function. */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }
  private emit(): void {
    this.listeners.forEach((fn) => fn());
  }

  /* ---------- accessors ---------- */

  getData(): CellMap {
    return this.data;
  }
  getTitle(): string {
    return this.title;
  }
  setTitle(t: string): void {
    this.title = t;
    this.emit();
  }
  /** Serializable state for persistence. */
  snapshot(): SheetSnapshot {
    return {
      data: this.data,
      title: this.title,
      rows: this._rows,
      cols: this._cols,
      colWidths: this.colWidths,
      gridlines: this.gridlines,
      merges: this.merges,
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
    };
  }
  /** Replace all state (no history entry) — used when loading from storage. */
  hydrate(snap: SheetSnapshot): void {
    this.data = snap.data || {};
    this.title = snap.title ?? this.title;
    this._rows = Math.min(MAX_ROWS, snap.rows ?? this.defaultRows);
    this._cols = Math.min(MAX_COLS, snap.cols ?? this.defaultCols);
    this.colWidths = { ...(snap.colWidths ?? {}) };
    this.gridlines = snap.gridlines ?? true;
    this.merges = (snap.merges ?? []).map((m) => ({ ...m }));
    this.frozenRows = snap.frozenRows ?? 0;
    this.frozenCols = snap.frozenCols ?? 0;
    this.cache = {};
    this.hist = [];
    this.future = [];
    this.emit();
  }
  /** Wipe everything back to an empty sheet. */
  reset(defaultTitle: string = DEFAULT_TITLE): void {
    this.data = {};
    this.title = defaultTitle;
    this._rows = this.defaultRows;
    this._cols = this.defaultCols;
    this.colWidths = {};
    this.gridlines = true;
    this.merges = [];
    this.frozenRows = 0;
    this.frozenCols = 0;
    this.cache = {};
    this.hist = [];
    this.future = [];
    this.emit();
  }

  /* ---------- frozen panes ---------- */
  getFrozenRows(): number {
    return Math.min(this.frozenRows, this._rows);
  }
  getFrozenCols(): number {
    return Math.min(this.frozenCols, this._cols);
  }
  setFreeze(rows: number, cols: number): void {
    const r = Math.max(0, Math.min(Math.round(rows), this._rows));
    const c = Math.max(0, Math.min(Math.round(cols), this._cols));
    if (r === this.frozenRows && c === this.frozenCols) return;
    this.checkpoint();
    this.frozenRows = r;
    this.frozenCols = c;
    this.emit();
  }

  /* ---------- dimensions (grow the sheet, like a real spreadsheet) ---------- */

  addRows(n = 1): void {
    const next = Math.min(MAX_ROWS, this._rows + Math.max(1, n));
    if (next === this._rows) return;
    this.checkpoint();
    this._rows = next;
    this.emit();
  }
  addCols(n = 1): void {
    const next = Math.min(MAX_COLS, this._cols + Math.max(1, n));
    if (next === this._cols) return;
    this.checkpoint();
    this._cols = next;
    this.emit();
  }

  /** Insert `n` rows before index `at`, shifting rows below down. Rewrites refs. */
  insertRows(at: number, n = 1): void {
    const count = Math.min(Math.max(1, n), MAX_ROWS - this._rows);
    if (count <= 0) return;
    const a = Math.max(0, Math.min(this._rows, at));
    this.checkpoint();
    const mapRow = (r: number) => (r < a ? r : r + count);
    const mapCol = (c: number) => c;
    this.applyRemap(mapRow, mapCol);
    this._rows += count;
    this.emit();
  }

  /** Insert `n` columns before index `at`, shifting columns right. Rewrites refs. */
  insertCols(at: number, n = 1): void {
    const count = Math.min(Math.max(1, n), MAX_COLS - this._cols);
    if (count <= 0) return;
    const a = Math.max(0, Math.min(this._cols, at));
    this.checkpoint();
    const mapRow = (r: number) => r;
    const mapCol = (c: number) => (c < a ? c : c + count);
    this.applyRemap(mapRow, mapCol);
    this._cols += count;
    this.emit();
  }

  /**
   * Delete rows r1..r2 (inclusive), shifting rows below up. References to deleted
   * rows become #REF!; references below shift. Always keeps at least one row.
   */
  deleteRows(r1: number, r2: number): void {
    const a = Math.max(0, Math.min(r1, r2));
    const b = Math.min(this._rows - 1, Math.max(r1, r2));
    const n = b - a + 1;
    if (n <= 0 || n >= this._rows) return;
    this.checkpoint();
    const mapRow = (r: number): number | null => (r < a ? r : r > b ? r - n : null);
    const mapCol = (c: number) => c;
    this.applyRemap(mapRow, mapCol);
    this._rows -= n;
    this.emit();
  }

  /**
   * Delete columns c1..c2 (inclusive), shifting columns left (widths included).
   * References to deleted columns become #REF!. Always keeps at least one column.
   */
  deleteCols(c1: number, c2: number): void {
    const a = Math.max(0, Math.min(c1, c2));
    const b = Math.min(this._cols - 1, Math.max(c1, c2));
    const n = b - a + 1;
    if (n <= 0 || n >= this._cols) return;
    this.checkpoint();
    const mapRow = (r: number) => r;
    const mapCol = (c: number): number | null => (c < a ? c : c > b ? c - n : null);
    this.applyRemap(mapRow, mapCol);
    this._cols -= n;
    this.emit();
  }

  /* ---------- structural remap (shared by insert/delete) ---------- */

  private rewriteCell(cell: CellData, mapRow: (r: number) => number | null, mapCol: (c: number) => number | null): CellData {
    if (cell.v && cell.v[0] === '=') {
      const nv = '=' + rewriteFormulaRefs(cell.v.slice(1), mapRow, mapCol);
      if (nv !== cell.v) return { ...cell, v: nv };
    }
    return cell;
  }

  private remapMerges(mapRow: (r: number) => number | null, mapCol: (c: number) => number | null): MergeRegion[] {
    const out: MergeRegion[] = [];
    for (const m of this.merges) {
      const rows: number[] = [];
      const cols: number[] = [];
      for (let r = m.r1; r <= m.r2; r++) {
        const nr = mapRow(r);
        if (nr !== null) rows.push(nr);
      }
      for (let c = m.c1; c <= m.c2; c++) {
        const nc = mapCol(c);
        if (nc !== null) cols.push(nc);
      }
      if (!rows.length || !cols.length) continue;
      const r1 = Math.min(...rows);
      const r2 = Math.max(...rows);
      const c1 = Math.min(...cols);
      const c2 = Math.max(...cols);
      if (r1 === r2 && c1 === c2) continue; // collapsed to a single cell
      out.push({ r1, c1, r2, c2 });
    }
    return out;
  }

  /** Reindex cells + column widths + merges and rewrite formulas for a remap. */
  private applyRemap(mapRow: (r: number) => number | null, mapCol: (c: number) => number | null): void {
    const data: CellMap = {};
    for (const k in this.data) {
      const comma = k.indexOf(',');
      const r = +k.slice(0, comma);
      const c = +k.slice(comma + 1);
      const nr = mapRow(r);
      const nc = mapCol(c);
      if (nr === null || nc === null) continue; // cell removed
      data[nr + ',' + nc] = this.rewriteCell(this.data[k], mapRow, mapCol);
    }
    const widths: Record<number, number> = {};
    for (const key in this.colWidths) {
      const nc = mapCol(+key);
      if (nc !== null) widths[nc] = this.colWidths[+key];
    }
    this.data = data;
    this.colWidths = widths;
    this.merges = this.remapMerges(mapRow, mapCol);
    this.cache = {};
  }

  /* ---------- column widths / gridlines (sheet metadata, persisted) ---------- */

  getColWidth(c: number): number {
    return this.colWidths[c] ?? this.defaultColWidth;
  }
  setColWidth(c: number, w: number): void {
    this.colWidths[c] = Math.max(MIN_COL_WIDTH, Math.min(MAX_COL_WIDTH, Math.round(w)));
    this.emit();
  }
  getGridlines(): boolean {
    return this.gridlines;
  }
  setGridlines(on: boolean): void {
    if (on === this.gridlines) return;
    this.checkpoint();
    this.gridlines = on;
    this.emit();
  }

  /* ---------- borders ---------- */

  /** Apply/clear custom borders over a range. See {@link BorderMode}. */
  applyBorders(range: Range, mode: BorderMode, color: string): void {
    this.mutate((d) => {
      forEachInRange(range, (r, c) => {
        const k = r + ',' + c;
        const cur = d[k] || {};
        if (mode === 'none') {
          const cell = { ...cur };
          delete cell.bd;
          d[k] = cell;
          return;
        }
        const bd: CellBorders = { ...cur.bd };
        const top = mode === 'all' || (mode === 'outer' && r === range.r1) || (mode === 'top' && r === range.r1);
        const bottom = mode === 'all' || (mode === 'outer' && r === range.r2) || (mode === 'bottom' && r === range.r2);
        const left = mode === 'all' || (mode === 'outer' && c === range.c1) || (mode === 'left' && c === range.c1);
        const right = mode === 'all' || (mode === 'outer' && c === range.c2) || (mode === 'right' && c === range.c2);
        if (top) bd.t = color;
        if (bottom) bd.b = color;
        if (left) bd.l = color;
        if (right) bd.r = color;
        const cell = { ...cur };
        if (bd.t || bd.r || bd.b || bd.l) cell.bd = bd;
        else delete cell.bd;
        d[k] = cell;
      });
    });
  }

  /* ---------- merges ---------- */

  getMerges(): MergeRegion[] {
    return this.merges;
  }
  /** The merge region covering (r,c), or null. */
  mergeAt(r: number, c: number): MergeRegion | null {
    for (const m of this.merges) {
      if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    }
    return null;
  }
  /** Grow a range so it fully contains every merge it touches (fixpoint). */
  expandRange(range: Range): Range {
    let { r1, r2, c1, c2 } = range;
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of this.merges) {
        if (m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1) {
          if (m.r1 < r1) (r1 = m.r1), (changed = true);
          if (m.r2 > r2) (r2 = m.r2), (changed = true);
          if (m.c1 < c1) (c1 = m.c1), (changed = true);
          if (m.c2 > c2) (c2 = m.c2), (changed = true);
        }
      }
    }
    return { r1, r2, c1, c2 };
  }
  /** True if any merge overlaps the range. */
  hasMergeIn(range: Range): boolean {
    return this.merges.some((m) => m.r1 <= range.r2 && m.r2 >= range.r1 && m.c1 <= range.c2 && m.c2 >= range.c1);
  }
  /** Merge a range into one cell (keeps the top-left content, clears the rest). */
  mergeCells(range: Range): void {
    const { r1, c1, r2, c2 } = range;
    if (r1 === r2 && c1 === c2) return;
    this.checkpoint();
    const data: CellMap = {};
    for (const k in this.data) data[k] = this.data[k];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r === r1 && c === c1) continue;
        delete data[r + ',' + c];
      }
    }
    this.data = data;
    this.merges = this.merges
      .filter((m) => !(m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1))
      .concat([{ r1, c1, r2, c2 }]);
    this.cache = {};
    this.emit();
  }
  /** Remove every merge that overlaps the range. */
  unmergeCells(range: Range): void {
    const kept = this.merges.filter((m) => !(m.r1 <= range.r2 && m.r2 >= range.r1 && m.c1 <= range.c2 && m.c2 >= range.c1));
    if (kept.length === this.merges.length) return;
    this.checkpoint();
    this.merges = kept;
    this.emit();
  }

  cellAt(r: number, c: number): CellData | undefined {
    return this.data[r + ',' + c];
  }
  /** Raw (unevaluated) cell input. */
  raw(r: number, c: number): string {
    const x = this.cellAt(r, c);
    return x && x.v != null ? x.v : '';
  }

  /* ---------- mutations (all funnel through mutate) ---------- */

  private mutate(fn: (d: CellMap) => void, takeSnapshot = true): void {
    if (takeSnapshot) this.checkpoint();
    const data: CellMap = {};
    for (const k in this.data) data[k] = { ...this.data[k] };
    fn(data);
    // prune cells that carry neither value nor formatting
    for (const k in data) {
      const c = data[k];
      const bd = c.bd;
      const noBorder = !bd || (!bd.t && !bd.r && !bd.b && !bd.l);
      if (!c.v && !c.b && !c.i && !c.s && !c.a && !c.f && !c.g && !c.co && !c.ff && noBorder) delete data[k];
    }
    this.cache = {};
    this.data = data;
    this.emit();
  }

  setRaw(r: number, c: number, v: string): void {
    this.mutate((d) => {
      const k = r + ',' + c;
      d[k] = { ...d[k], v };
    });
  }

  setAttr<K extends keyof CellData>(range: Range, attr: K, value: CellData[K]): void {
    this.mutate((d) => {
      forEachInRange(range, (r, c) => {
        const k = r + ',' + c;
        d[k] = { ...d[k], [attr]: value } as CellData;
      });
    });
  }

  clearRange(range: Range): void {
    this.mutate((d) => {
      forEachInRange(range, (r, c) => {
        delete d[r + ',' + c];
      });
    });
  }

  /** Paste a 2D grid of raw values with its top-left at `anchor`. */
  paste(anchor: { r: number; c: number }, grid: string[][]): void {
    this.mutate((d) => {
      grid.forEach((row, dr) =>
        row.forEach((v, dc) => {
          const r = anchor.r + dr;
          const c = anchor.c + dc;
          if (r < this.rows && c < this.cols) d[r + ',' + c] = { ...d[r + ',' + c], v };
        }),
      );
    });
  }

  /* ---------- history ---------- */

  get canUndo(): boolean {
    return this.hist.length > 0;
  }
  get canRedo(): boolean {
    return this.future.length > 0;
  }
  undo(): void {
    if (!this.hist.length) return;
    const prev = this.hist[this.hist.length - 1];
    this.hist = this.hist.slice(0, -1);
    this.future = this.future.concat([this.capture()]);
    this.restore(prev);
    this.emit();
  }
  redo(): void {
    if (!this.future.length) return;
    const next = this.future[this.future.length - 1];
    this.future = this.future.slice(0, -1);
    this.hist = this.hist.concat([this.capture()]);
    this.restore(next);
    this.emit();
  }

  /* ---------- evaluation ---------- */

  /**
   * Evaluate a cell. Numbers/strings/booleans come back typed; formula errors
   * come back as strings like "#DIV/0!". May throw {e} for cycle detection —
   * callers rendering a single cell should use {@link safeValue} instead.
   */
  value(r: number, c: number, seen: Set<string> = new Set()): CellValue {
    if (r < 0 || c < 0 || r >= this.rows || c >= this.cols) return '';
    const k = r + ',' + c;
    const rw = this.raw(r, c);
    if (rw === '') return '';
    if (rw[0] !== '=') {
      const n = numeric(rw);
      return n === null ? rw : n;
    }
    if (this.cache[k] !== undefined) return this.cache[k];
    if (seen.has(k)) throw { e: '#CYCLE!' } as FormulaError;
    seen.add(k);
    let out: CellValue;
    try {
      out = evaluate(rw.slice(1), (rr, cc) => this.value(rr, cc, seen));
    } catch (err) {
      out = err && (err as FormulaError).e ? (err as FormulaError).e : '#ERROR!';
    }
    seen.delete(k);
    if (typeof out === 'number' && !isFinite(out)) out = '#DIV/0!';
    this.cache[k] = out;
    return out;
  }

  /** Evaluate a cell, converting cycle errors into a value instead of throwing. */
  safeValue(r: number, c: number): CellValue {
    try {
      return this.value(r, c, new Set());
    } catch (err) {
      return (err as FormulaError)?.e || '#ERROR!';
    }
  }

  /** Sum / numeric count / non-empty count over a rectangle (for the status bar). */
  stats(range: Range): RangeStats {
    let sum = 0;
    let count = 0;
    let filled = 0;
    forEachInRange(range, (r, c) => {
      const v = this.safeValue(r, c);
      if (v !== '' && v != null) filled++;
      if (typeof v === 'number') {
        sum += v;
        count++;
      }
    });
    return { sum, count, filled };
  }

  /* ---------- static helpers ---------- */
  static colName = colName;
  static colIndex = colIndex;
}
