// Shared plugin state used across modules

export const SCAN_DEBOUNCE_MS = 500;

let _lastScanTime = 0;
export function getLastScanTime(): number { return _lastScanTime; }
export function setLastScanTime(ts: number): void { _lastScanTime = ts; }

let _isCreatingTable = false;
export function getIsCreatingTable(): boolean { return _isCreatingTable; }
export function setIsCreatingTable(v: boolean): void { _isCreatingTable = v; }

// Map for temporary instance management if needed by router
export const cellInstanceMap: Map<string, InstanceNode> = new Map();

// Track last scan result table id (optional future use)
let _currentTableId: string | undefined;
export function getCurrentTableId(): string | undefined { return _currentTableId; }
export function setCurrentTableId(id: string | undefined): void { _currentTableId = id; }


