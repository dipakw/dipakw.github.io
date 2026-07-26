import { SpreadsheetEngine } from './SpreadsheetEngine';
import type { EngineOptions } from './SpreadsheetEngine';
import type { SheetSnapshot } from './types';

/** One sheet inside a persisted workbook (its tab name lives in `sheet.title`). */
export interface WorkbookSheet {
  id: string;
  sheet: SheetSnapshot;
}

/** Serializable workbook — what a StorageAdapter persists. */
export interface WorkbookSnapshot {
  /** Workbook/document name (shown in the title bar). */
  name: string;
  /** Id of the active sheet. */
  activeId: string;
  sheets: WorkbookSheet[];
}

export interface WorkbookOptions {
  rows?: number;
  cols?: number;
  defaultColWidth?: number;
  /** Default workbook name. */
  defaultName?: string;
  /** Initial state (accepts a legacy single-sheet snapshot too). */
  snapshot?: WorkbookSnapshot | SheetSnapshot | null;
}

interface LiveSheet {
  id: string;
  engine: SpreadsheetEngine;
  unsub: () => void;
}

const DEFAULT_NAME = 'Untitled workbook';

/** Hard cap on the number of sheets in a workbook. */
export const MAX_SHEETS = 20;

function newId(): string {
  return 'sh_' + Math.random().toString(36).slice(2, 9);
}

/** Coerce whatever was persisted into a WorkbookSnapshot (migrating legacy sheets). */
function toWorkbook(snap: WorkbookSnapshot | SheetSnapshot | null | undefined): WorkbookSnapshot | null {
  if (!snap) return null;
  const anySnap = snap as unknown as Record<string, unknown>;
  if (Array.isArray(anySnap.sheets)) return snap as WorkbookSnapshot;
  // Legacy single-sheet snapshot → wrap it into a one-sheet workbook.
  if (anySnap.data || anySnap.title != null) {
    const sheet = snap as SheetSnapshot;
    const id = newId();
    return { name: sheet.title || DEFAULT_NAME, activeId: id, sheets: [{ id, sheet }] };
  }
  return null;
}

/**
 * A collection of sheets, each backed by its own {@link SpreadsheetEngine}.
 * Framework-agnostic. It forwards every sheet's changes through a single
 * `subscribe`, so a UI only needs to watch the workbook to stay in sync with
 * edits on the active sheet *and* structural changes (add/rename/remove/switch).
 */
export class Workbook {
  private opts: WorkbookOptions;
  private name: string = DEFAULT_NAME;
  private sheets: LiveSheet[] = [];
  private activeId = '';
  private listeners = new Set<() => void>();

  constructor(opts: WorkbookOptions = {}) {
    this.opts = opts;
    this.load(opts.snapshot ?? null);
  }

  /* ---------- subscription ---------- */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }
  private emit(): void {
    this.listeners.forEach((fn) => fn());
  }

  /* ---------- internals ---------- */
  private engineOpts(snapshot: SheetSnapshot | null, fallbackTitle: string): EngineOptions {
    return {
      rows: this.opts.rows,
      cols: this.opts.cols,
      defaultColWidth: this.opts.defaultColWidth,
      defaultTitle: snapshot?.title ?? fallbackTitle,
      snapshot,
    };
  }
  private makeSheet(id: string, snapshot: SheetSnapshot | null, fallbackTitle: string): LiveSheet {
    const engine = new SpreadsheetEngine(this.engineOpts(snapshot, fallbackTitle));
    const unsub = engine.subscribe(() => this.emit());
    return { id, engine, unsub };
  }
  private load(snap: WorkbookSnapshot | SheetSnapshot | null): void {
    this.sheets.forEach((s) => s.unsub());
    const wb = toWorkbook(snap);
    if (wb && wb.sheets.length) {
      this.name = wb.name || this.opts.defaultName || DEFAULT_NAME;
      this.sheets = wb.sheets.map((s, i) => this.makeSheet(s.id, s.sheet, 'Sheet ' + (i + 1)));
      this.activeId = wb.sheets.some((s) => s.id === wb.activeId) ? wb.activeId : this.sheets[0].id;
    } else {
      this.name = this.opts.defaultName || DEFAULT_NAME;
      const first = this.makeSheet(newId(), null, 'Sheet 1');
      this.sheets = [first];
      this.activeId = first.id;
    }
  }
  private uniqueName(): string {
    const names = new Set(this.sheets.map((s) => s.engine.getTitle()));
    let n = this.sheets.length + 1;
    while (names.has('Sheet ' + n)) n++;
    return 'Sheet ' + n;
  }

  /* ---------- name ---------- */
  getName(): string {
    return this.name;
  }
  setName(name: string): void {
    this.name = name;
    this.emit();
  }

  /* ---------- sheets ---------- */
  /** Tab list: id + display name (the sheet's title). */
  list(): Array<{ id: string; name: string }> {
    return this.sheets.map((s) => ({ id: s.id, name: s.engine.getTitle() }));
  }
  count(): number {
    return this.sheets.length;
  }
  getActiveId(): string {
    return this.activeId;
  }
  getActive(): SpreadsheetEngine {
    return (this.sheets.find((s) => s.id === this.activeId) || this.sheets[0]).engine;
  }
  setActive(id: string): void {
    if (id !== this.activeId && this.sheets.some((s) => s.id === id)) {
      this.activeId = id;
      this.emit();
    }
  }
  addSheet(): string {
    if (this.sheets.length >= MAX_SHEETS) return this.activeId;
    const s = this.makeSheet(newId(), null, this.uniqueName());
    this.sheets.push(s);
    this.activeId = s.id;
    this.emit();
    return s.id;
  }
  renameSheet(id: string, name: string): void {
    const s = this.sheets.find((x) => x.id === id);
    if (s) s.engine.setTitle(name.trim() || s.engine.getTitle());
  }
  removeSheet(id: string): void {
    if (this.sheets.length <= 1) return;
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx < 0) return;
    this.sheets[idx].unsub();
    this.sheets.splice(idx, 1);
    if (this.activeId === id) this.activeId = this.sheets[Math.max(0, idx - 1)].id;
    this.emit();
  }

  /* ---------- persistence ---------- */
  snapshot(): WorkbookSnapshot {
    return {
      name: this.name,
      activeId: this.activeId,
      sheets: this.sheets.map((s) => ({ id: s.id, sheet: s.engine.snapshot() })),
    };
  }
  /** Replace all state (from storage). */
  hydrate(snap: WorkbookSnapshot | SheetSnapshot | null): void {
    this.load(snap);
    this.emit();
  }
  /** Reset to a single empty sheet. */
  reset(): void {
    this.load(null);
    this.emit();
  }
}
