import { createBlankGrid } from './lib/grid';
import { GRID_COLUMNS, type GridCellValue, type PresetType } from './types';

const JOURNAL_TEMPLATE_HEADERS = ['#', 'Account Name', 'Debit', 'Credit'] as const;
const JOURNAL_COLUMN_WIDTHS = [32, 220, 84, 84, 74, 74, 76] as const;
const JOURNAL_TOTAL_WIDTH = JOURNAL_COLUMN_WIDTHS.reduce((total, width) => total + width, 0);
const T_ACCOUNT_COLUMN_WIDTH = JOURNAL_TOTAL_WIDTH / GRID_COLUMNS;
const T_ACCOUNT_COLUMN_WIDTHS = Array.from({ length: GRID_COLUMNS }, () => T_ACCOUNT_COLUMN_WIDTH);
const BLANK_COLUMN_WIDTH = 72;
const BLANK_COLUMN_WIDTHS = Array.from({ length: GRID_COLUMNS }, () => BLANK_COLUMN_WIDTH);

const T_ACCOUNT_BLOCKS = [
  { titleRow: 2, topRow: 3, entryStartRow: 4, entryEndRow: 10, totalRow: 11, title: 'Account 1' },
  { titleRow: 13, topRow: 14, entryStartRow: 15, entryEndRow: 21, totalRow: 22, title: 'Account 2' },
] as const;
const T_ACCOUNT_DEBIT_COL = 1;
const T_ACCOUNT_CREDIT_COL = 2;

export function buildPresetGrid(preset: PresetType): GridCellValue[][] {
  const next = createBlankGrid();

  if (preset === 'journal_entry') {
    JOURNAL_TEMPLATE_HEADERS.forEach((header, columnIndex) => {
      next[0][columnIndex] = header;
    });
    next[1][0] = 1;
    return next;
  }

  if (preset === 't_account') {
    T_ACCOUNT_BLOCKS.forEach((block) => {
      next[block.titleRow][T_ACCOUNT_DEBIT_COL] = block.title;
      next[block.topRow][T_ACCOUNT_DEBIT_COL] = 'Dr';
      next[block.topRow][T_ACCOUNT_CREDIT_COL] = 'Cr';

      const debitRange = `${toA1Col(T_ACCOUNT_DEBIT_COL)}${block.entryStartRow + 1}:${toA1Col(T_ACCOUNT_DEBIT_COL)}${block.entryEndRow + 1}`;
      const creditRange = `${toA1Col(T_ACCOUNT_CREDIT_COL)}${block.entryStartRow + 1}:${toA1Col(T_ACCOUNT_CREDIT_COL)}${block.entryEndRow + 1}`;
      const netExpression = `SUM(${debitRange})-SUM(${creditRange})`;
      const hasEntriesExpression = `COUNTA(${debitRange})+COUNTA(${creditRange})>0`;

      next[block.totalRow][T_ACCOUNT_DEBIT_COL] =
        `=IF(${hasEntriesExpression},IF(${netExpression}>=0,${netExpression},""),"")`;
      next[block.totalRow][T_ACCOUNT_CREDIT_COL] =
        `=IF(${hasEntriesExpression},IF(${netExpression}<0,ABS(${netExpression}),""),"")`;
    });
  }

  return next;
}

export function getPresetColumnCount(_preset: PresetType): number {
  return GRID_COLUMNS;
}

export function getPresetColumnHeaders(_preset: PresetType): string[] {
  return Array.from({ length: GRID_COLUMNS }, (_, index) => String.fromCharCode(65 + index));
}

export function getPresetHiddenColumns(_preset: PresetType): number[] {
  return [];
}

export function getPresetColumnWidths(preset: PresetType): number[] {
  if (preset === 'journal_entry') {
    return [...JOURNAL_COLUMN_WIDTHS];
  }

  if (preset === 'blank') {
    return [...BLANK_COLUMN_WIDTHS];
  }

  return [...T_ACCOUNT_COLUMN_WIDTHS];
}

export function isTAccountTopBorderCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some(
    (block) => row === block.topRow && (col === T_ACCOUNT_DEBIT_COL || col === T_ACCOUNT_CREDIT_COL),
  );
}

export function isTAccountDividerCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some(
    (block) => col === T_ACCOUNT_DEBIT_COL && row >= block.topRow && row <= block.totalRow,
  );
}

export function isTAccountTitleCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some((block) => row === block.titleRow && col === T_ACCOUNT_DEBIT_COL);
}

export function isTAccountLabelCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some(
    (block) => row === block.topRow && (col === T_ACCOUNT_DEBIT_COL || col === T_ACCOUNT_CREDIT_COL),
  );
}

export function isTAccountSumBorderCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some(
    (block) => row === block.totalRow && (col === T_ACCOUNT_DEBIT_COL || col === T_ACCOUNT_CREDIT_COL),
  );
}

export function isTAccountTotalCell(row: number, col: number): boolean {
  return T_ACCOUNT_BLOCKS.some(
    (block) => row === block.totalRow && (col === T_ACCOUNT_DEBIT_COL || col === T_ACCOUNT_CREDIT_COL),
  );
}

export function getPresetMergeCells(preset: PresetType): Array<{
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
}> {
  if (preset !== 't_account') {
    return [];
  }

  return T_ACCOUNT_BLOCKS.map((block) => ({
    row: block.titleRow,
    col: T_ACCOUNT_DEBIT_COL,
    rowspan: 1,
    colspan: 2,
  }));
}

function toA1Col(columnIndex: number): string {
  return String.fromCharCode(65 + columnIndex);
}
