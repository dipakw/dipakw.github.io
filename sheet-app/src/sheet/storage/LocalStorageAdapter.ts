import type { WorkbookSnapshot } from '../engine/Workbook';
import type { StorageAdapter } from './StorageAdapter';

/**
 * Default storage adapter: persists the whole workbook as one JSON blob in
 * localStorage. Bundled with the package as a batteries-included default — the
 * component only depends on the {@link StorageAdapter} port, so you can swap in
 * an IndexedDB / REST / collaborative adapter without touching the engine or UI.
 */
export class LocalStorageAdapter implements StorageAdapter {
  readonly label = 'saved locally';

  constructor(private readonly key: string = 'sheet.workbook.v1') {}

  load(): WorkbookSnapshot | null {
    try {
      return JSON.parse(localStorage.getItem(this.key) || 'null');
    } catch {
      return null;
    }
  }

  save(snapshot: WorkbookSnapshot): void {
    try {
      localStorage.setItem(this.key, JSON.stringify(snapshot));
    } catch {
      /* ignore quota / private-mode errors */
    }
  }

  clear(): void {
    try {
      localStorage.removeItem(this.key);
    } catch {
      /* ignore */
    }
  }
}
