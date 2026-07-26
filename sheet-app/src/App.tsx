import { useMemo } from 'react';
import { Sheet, LocalStorageAdapter } from './sheet';

/**
 * App shell. The Sheet package is self-contained (component, engine, styles,
 * storage). Here we just pick a storage adapter and mount it — swap the adapter
 * to change where data persists. The key preserves any existing local data.
 */
export default function App() {
  const storage = useMemo(() => new LocalStorageAdapter('ledger.sheet.v2'), []);
  return <Sheet storage={storage} />;
}
