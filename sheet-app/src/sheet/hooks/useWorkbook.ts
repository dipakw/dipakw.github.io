import { useEffect, useReducer, useRef, useState } from 'react';
import type { Workbook } from '../engine/Workbook';
import type { StorageAdapter } from '../storage/StorageAdapter';

export interface UseWorkbookOptions {
  autosaveMs?: number;
}

export interface UseWorkbookResult {
  savedLabel: string;
  hydrated: boolean;
}

/**
 * Wires a {@link Workbook} to a {@link StorageAdapter} inside React:
 *  - re-renders on any change (active-sheet edits *and* structural changes),
 *  - hydrates from storage once on mount,
 *  - autosaves (debounced) the whole workbook on every change.
 *
 * The workbook forwards every sheet's changes through its own `subscribe`, so
 * one subscription here covers everything.
 */
export function useWorkbook(workbook: Workbook, adapter: StorageAdapter, opts: UseWorkbookOptions = {}): UseWorkbookResult {
  const autosaveMs = opts.autosaveMs ?? 250;
  const [, force] = useReducer((x: number) => x + 1, 0);
  const [savedLabel, setSavedLabel] = useState('');
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => workbook.subscribe(force), [workbook]);

  useEffect(() => {
    let alive = true;
    Promise.resolve(adapter.load()).then((snap) => {
      if (!alive) return;
      if (snap) workbook.hydrate(snap);
      setHydrated(true);
    });
    return () => {
      alive = false;
    };
  }, [workbook, adapter]);

  const timer = useRef<ReturnType<typeof setTimeout>>();
  useEffect(() => {
    const unsub = workbook.subscribe(() => {
      clearTimeout(timer.current);
      timer.current = setTimeout(() => {
        Promise.resolve(adapter.save(workbook.snapshot())).then(() => {
          setSavedLabel(adapter.label ?? 'saved');
        });
      }, autosaveMs);
    });
    return () => {
      clearTimeout(timer.current);
      unsub();
    };
  }, [workbook, adapter, autosaveMs]);

  return { savedLabel, hydrated };
}
