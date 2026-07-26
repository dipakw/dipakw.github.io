// Public surface of the reusable Sheet package.
export { Sheet } from './Sheet';
export type { SheetProps } from './Sheet';

export { SpreadsheetEngine } from './engine/SpreadsheetEngine';
export type { EngineOptions, RangeStats } from './engine/SpreadsheetEngine';
export { Workbook } from './engine/Workbook';
export type { WorkbookOptions, WorkbookSnapshot, WorkbookSheet } from './engine/Workbook';
export { colName, colIndex, numeric } from './engine/cellref';
export { format } from './engine/format';
export * from './engine/types';

// Storage: the package owns the *port* (interface) and ships a batteries-included
// localStorage adapter. Swap in your own adapter to change where data persists.
export type { StorageAdapter } from './storage/StorageAdapter';
export { LocalStorageAdapter } from './storage/LocalStorageAdapter';

export { useWorkbook } from './hooks/useWorkbook';
export type { UseWorkbookOptions, UseWorkbookResult } from './hooks/useWorkbook';
export { useTheme } from './hooks/useTheme';
export type { Theme } from './hooks/useTheme';
