import { GRID_COLUMNS, GRID_ROWS, type GridCellValue } from '../types';

export function createBlankGrid(): GridCellValue[][] {
  return Array.from({ length: GRID_ROWS }, () =>
    Array.from({ length: GRID_COLUMNS }, () => null),
  );
}

export function hasNonEmptyCells(matrix: GridCellValue[][]): boolean {
  return matrix.some((row) => row.some((cell) => !isBlank(cell)));
}

export function normalizeDisplayValue(value: unknown): GridCellValue {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length === 0 ? null : value;
  }

  return String(value);
}

export function isFormula(rawValue: unknown): rawValue is string {
  return typeof rawValue === 'string' && rawValue.trim().startsWith('=');
}

function isBlank(value: GridCellValue): boolean {
  return value === null || (typeof value === 'string' && value.trim().length === 0);
}
