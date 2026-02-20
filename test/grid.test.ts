import { describe, expect, it } from 'vitest';
import { GRID_COLUMNS, GRID_ROWS } from '../src/types';
import { createBlankGrid, hasNonEmptyCells, isFormula, normalizeDisplayValue } from '../src/lib/grid';

describe('grid utilities', () => {
  it('creates a blank GRID_ROWS x GRID_COLUMNS matrix', () => {
    const matrix = createBlankGrid();

    expect(matrix).toHaveLength(GRID_ROWS);
    expect(matrix.every((row) => row.length === GRID_COLUMNS)).toBe(true);
    expect(matrix.flat().every((value) => value === null)).toBe(true);
  });

  it('detects non-empty cells', () => {
    expect(hasNonEmptyCells(createBlankGrid())).toBe(false);
    expect(hasNonEmptyCells([[null, '   '], [null, null]])).toBe(false);
    expect(hasNonEmptyCells([[null, 'x'], [null, null]])).toBe(true);
    expect(hasNonEmptyCells([[null, 0], [null, null]])).toBe(true);
    expect(hasNonEmptyCells([[null, false], [null, null]])).toBe(true);
  });

  it('normalizes display values consistently', () => {
    expect(normalizeDisplayValue(null)).toBeNull();
    expect(normalizeDisplayValue(undefined)).toBeNull();
    expect(normalizeDisplayValue('')).toBeNull();
    expect(normalizeDisplayValue('   ')).toBeNull();
    expect(normalizeDisplayValue('  text  ')).toBe('  text  ');
    expect(normalizeDisplayValue(10)).toBe(10);
    expect(normalizeDisplayValue(false)).toBe(false);
    expect(normalizeDisplayValue({ x: 1 })).toBe('[object Object]');
  });

  it('identifies formula-like strings', () => {
    expect(isFormula('=SUM(A1:A2)')).toBe(true);
    expect(isFormula('   =A1+1')).toBe(true);
    expect(isFormula('A1+1')).toBe(false);
    expect(isFormula(3)).toBe(false);
    expect(isFormula(null)).toBe(false);
  });
});
