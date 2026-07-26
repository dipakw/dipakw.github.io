import type { WorkbookSnapshot } from '../engine/Workbook';

/**
 * Pluggable persistence contract for a workbook.
 *
 * Every method may be sync or async so the same interface works for
 * localStorage today and, later, IndexedDB / a REST or GraphQL backend / a
 * collaborative store — swap the adapter, the engine and UI stay unchanged.
 */
export interface StorageAdapter {
  /** Human-readable "saved" indicator (e.g. "saved locally", "synced"). */
  readonly label?: string;

  /** Load the persisted snapshot, or null if nothing is stored. */
  load(): WorkbookSnapshot | null | Promise<WorkbookSnapshot | null>;

  /** Persist the current snapshot. Called debounced on change. */
  save(snapshot: WorkbookSnapshot): void | Promise<void>;

  /** Remove all persisted data (used by "Reset all"). Optional. */
  clear?(): void | Promise<void>;
}
