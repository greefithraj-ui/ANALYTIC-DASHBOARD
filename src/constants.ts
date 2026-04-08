
import { ColumnMapping, SheetConfig } from './types';

export const DEFAULT_MAPPING: ColumnMapping = {
  uid: 'UID',
  ringStatus: 'RING STATUS',
  date: 'DATE',
  batchNo: 'BATCH NO',
  inward: 'INWARD',
  sku: 'SKU',
  reason: 'REASON',
  quantity: 'QTY',
  movedToInventory: 'MOVED TO INVENTORY'
};

export const INITIAL_CONFIG: SheetConfig = {
  url: '',
  sheetName: 'Sheet1',
  range: 'A1:Z1000',
  mapping: { ...DEFAULT_MAPPING }
};

export const COLORS = {
  metallicGreen: '#10b981',
  metallicRed: '#ef4444',
  slate400: '#94a3b8',
  slate800: '#1e293b'
};
