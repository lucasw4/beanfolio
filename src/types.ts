export const GRID_COLUMNS = 12;
export const GRID_ROWS = 120;

export type GridCellValue = string | number | boolean | null;

export interface CellSnapshot {
  displayValue: GridCellValue;
  formula?: string;
}

export interface SaveRequest {
  label?: string;
  cells: CellSnapshot[][];
  columnCount: number;
}

export interface LedgerTarget {
  spreadsheetId: string;
  spreadsheetTitle: string;
  sheetId: number;
  sheetTitle: string;
}

export type PresetType = 'blank' | 'journal_entry' | 't_account';
